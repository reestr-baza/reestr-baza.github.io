import { buildLookupIndex } from '../formula/functions';
import { evaluateToScalar, type EvalCtx, type EvalHost } from '../formula/evaluate';
import { parseFormula, type Node } from '../formula/parse';
import { adjustFormula, remapColumns, renameSheetRefs, shiftFormula, type StructuralChange } from '../formula/refs';
import { compareScalars, parseLooseNumber } from '../formula/coerce';
import { ERR, isErr, type RangeVal, type Scalar } from '../formula/values';
import { uid } from '../lib/ids';
import { formatScalar, isUrl } from './format';
import {
  DEFAULT_COL_WIDTH,
  DENSITY_HEIGHT,
  type Cell,
  type CellStyle,
  type ColId,
  type Column,
  type FilterSpec,
  type Merge,
  type NamedValue,
  type Row,
  type RowId,
  type Sheet,
  type SheetId,
  type SheetMeta,
  type WorkbookMeta,
} from './types';

// ─── транзакции ──────────────────────────────────────────────────────────────

type Patch =
  | { k: 'row'; sheetId: SheetId; rowId: RowId; prev: Row | undefined; next: Row | undefined }
  | { k: 'order'; sheetId: SheetId; prev: RowId[]; next: RowId[] }
  | { k: 'cols'; sheetId: SheetId; prev: Column[]; next: Column[] }
  | { k: 'sheet'; sheetId: SheetId; prev: Partial<SheetMeta>; next: Partial<SheetMeta> }
  | { k: 'addSheet'; sheet: Sheet; index: number }
  | { k: 'removeSheet'; sheet: Sheet; index: number }
  | { k: 'wb'; prev: Partial<WorkbookMeta>; next: Partial<WorkbookMeta> };

interface Tx {
  label: string;
  /** Сделано в режиме просмотра (фильтр, сортировка) — и отменять его можно там же */
  view?: boolean;
  patches: Patch[];
  rowPatch: Map<string, Extract<Patch, { k: 'row' }>>;
  /** Выделение до изменения — чтобы отмена возвращала курсор на место. */
  selection?: unknown;
}

export interface CommitInfo {
  rows: { sheetId: SheetId; rowId: RowId; row: Row | undefined }[];
  sheets: Set<SheetId>;
  removedSheets: SheetId[];
  wb: boolean;
  structural: boolean;
}

export interface CellDisplay {
  text: string;
  value: Scalar;
  style: CellStyle;
  img?: string;
  href?: string;
  isNumber: boolean;
  isError: boolean;
  isFormula: boolean;
  note?: string;
}

export interface SheetView {
  /** Видимые строки по порядку */
  rows: RowId[];
  /** Физический индекс строки (для номера и формул) */
  rowIndex: Map<RowId, number>;
  /** Индексы видимых столбцов */
  cols: number[];
  /** Видимые строки как множество — для ПРОМЕЖУТОЧНЫЕ.ИТОГИ */
  visible: Set<RowId>;
  filtered: boolean;
  totalRows: number;
}

type Listener = () => void;

const HIDDEN_STYLE: CellStyle = {};

export function emptySheet(name: string, columns: string[] = ['A', 'B', 'C', 'D', 'E', 'F'], rows = 50): Sheet {
  const cols: Column[] = columns.map((n) => ({ id: uid(6), name: n, w: DEFAULT_COL_WIDTH }));
  const map = new Map<RowId, Row>();
  const order: RowId[] = [];
  for (let i = 0; i < rows; i++) {
    const id = uid(8);
    order.push(id);
    map.set(id, { id, cells: {} });
  }
  return {
    id: uid(8),
    name,
    columns: cols,
    rowOrder: order,
    rows: map,
    keyColId: cols[0]?.id,
    frozen: 1,
    density: 'S',
    filters: {},
  };
}

export class Store implements EvalHost {
  meta: WorkbookMeta;
  sheets = new Map<SheetId, Sheet>();

  /** Меняется при любом изменении — повод перерисоваться. */
  version = 0;
  /** Меняется при изменении данных — сброс кэша вычислений. */
  calcEpoch = 0;
  /** Меняется при изменении структуры/размеров — пересчёт геометрии сетки. */
  layoutVersion = 0;

  private listeners = new Set<Listener>();
  private commitListeners = new Set<(c: CommitInfo) => void>();
  private undoStack: Tx[] = [];
  private redoStack: Tx[] = [];
  private tx: Tx | null = null;
  private txDepth = 0;

  // вычисления
  private compiled = new Map<string, Node | Error>();
  private calcCache = new Map<string, Scalar>();
  private calcStack = new Set<string>();
  private lookupCache = new Map<string, Map<string, number>>();
  private colIndex = new WeakMap<Column[], Map<ColId, number>>();

  // представления (фильтры)
  private views = new Map<SheetId, { key: string; view: SheetView }>();
  private pinned = new Map<SheetId, Set<RowId>>();
  private filterEpoch = 0;
  search = '';

  /** Сохраняет/восстанавливает выделение UI при отмене. Назначается интерфейсом. */
  selectionProvider: { get(): unknown; set(v: unknown): void } | null = null;

  /** Режим просмотра: изменения данных отклоняются (кроме фильтров и сортировки). */
  readOnly = false;
  /** Вызывается, когда изменение отклонено режимом просмотра. */
  onBlocked: (() => void) | null = null;
  private allowDepth = 0;

  constructor(meta: WorkbookMeta, sheets: Sheet[]) {
    this.meta = meta;
    for (const s of sheets) this.sheets.set(s.id, s);
  }

  // ─── подписки ──────────────────────────────────────────────────────────────

  subscribe = (fn: Listener) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  onCommit(fn: (c: CommitInfo) => void) {
    this.commitListeners.add(fn);
    return () => this.commitListeners.delete(fn);
  }

  getVersion = () => this.version;

  private emit() {
    this.version++;
    for (const l of this.listeners) l();
  }

  // ─── доступ ────────────────────────────────────────────────────────────────

  get activeSheet(): Sheet {
    return this.sheets.get(this.meta.activeSheet) ?? this.sheets.values().next().value!;
  }

  sheetList(): Sheet[] {
    return this.meta.sheetIds.map((id) => this.sheets.get(id)!).filter(Boolean);
  }

  sheet(id: SheetId): Sheet {
    const s = this.sheets.get(id);
    if (!s) throw new Error('Лист не найден');
    return s;
  }

  colIndexOf(sheet: Sheet, colId: ColId): number {
    let m = this.colIndex.get(sheet.columns);
    if (!m) {
      m = new Map(sheet.columns.map((c, i) => [c.id, i]));
      this.colIndex.set(sheet.columns, m);
    }
    return m.get(colId) ?? -1;
  }

  rowHeight(sheet: Sheet, row: Row | undefined): number {
    return row?.h ?? DENSITY_HEIGHT[sheet.density];
  }

  canUndo() {
    return this.undoStack.length > 0;
  }
  canRedo() {
    return this.redoStack.length > 0;
  }
  undoLabel() {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }
  redoLabel() {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  // ─── транзакции ────────────────────────────────────────────────────────────

  /** Выполнить действие, разрешённое и в режиме просмотра (фильтр, сортировка). */
  allow<T>(fn: () => T): T {
    this.allowDepth++;
    try {
      return fn();
    } finally {
      this.allowDepth--;
    }
  }

  /** Отклонено ли изменение режимом просмотра — страховка для всех путей изменения данных. */
  private blocked(): boolean {
    if (!this.readOnly || this.allowDepth > 0) return false;
    this.onBlocked?.();
    return true;
  }

  transact<T>(label: string, fn: () => T): T {
    if (this.txDepth === 0 && this.blocked()) return undefined as T;
    if (this.txDepth === 0) {
      this.tx = { label, patches: [], rowPatch: new Map(), selection: this.selectionProvider?.get(), view: this.allowDepth > 0 };
    }
    this.txDepth++;
    try {
      return fn();
    } finally {
      this.txDepth--;
      if (this.txDepth === 0) {
        const tx = this.tx!;
        this.tx = null;
        if (tx.patches.length) {
          this.undoStack.push(tx);
          if (this.undoStack.length > 200) this.undoStack.shift();
          this.redoStack = [];
          this.afterChange(tx.patches, 'forward');
        }
      }
    }
  }

  private record(p: Patch) {
    if (!this.tx) throw new Error('Изменение вне транзакции');
    if (p.k === 'row') {
      const key = p.sheetId + '/' + p.rowId;
      const existing = this.tx.rowPatch.get(key);
      if (existing) {
        existing.next = p.next;
        return;
      }
      this.tx.rowPatch.set(key, p);
    }
    this.tx.patches.push(p);
  }

  undo() {
    const top = this.undoStack[this.undoStack.length - 1];
    if (!top || (!top.view && this.blocked())) return;
    const tx = this.undoStack.pop();
    if (!tx) return;
    const redoSel = this.selectionProvider?.get();
    for (let i = tx.patches.length - 1; i >= 0; i--) this.applyPatch(tx.patches[i], 'backward');
    this.redoStack.push({ ...tx, selection: redoSel });
    this.afterChange(tx.patches, 'backward');
    if (tx.selection !== undefined) this.selectionProvider?.set(tx.selection);
  }

  redo() {
    const top = this.redoStack[this.redoStack.length - 1];
    if (!top || (!top.view && this.blocked())) return;
    const tx = this.redoStack.pop();
    if (!tx) return;
    const sel = this.selectionProvider?.get();
    for (const p of tx.patches) this.applyPatch(p, 'forward');
    this.undoStack.push({ ...tx, selection: sel });
    this.afterChange(tx.patches, 'forward');
    if (tx.selection !== undefined) this.selectionProvider?.set(tx.selection);
  }

  private applyPatch(p: Patch, dir: 'forward' | 'backward') {
    const val = <T>(prev: T, next: T) => (dir === 'forward' ? next : prev);
    switch (p.k) {
      case 'row': {
        const s = this.sheets.get(p.sheetId);
        if (!s) return;
        const r = val(p.prev, p.next);
        if (r) s.rows.set(p.rowId, r);
        else s.rows.delete(p.rowId);
        return;
      }
      case 'order': {
        const s = this.sheets.get(p.sheetId);
        if (s) s.rowOrder = val(p.prev, p.next);
        return;
      }
      case 'cols': {
        const s = this.sheets.get(p.sheetId);
        if (s) s.columns = val(p.prev, p.next);
        return;
      }
      case 'sheet': {
        const s = this.sheets.get(p.sheetId);
        if (s) Object.assign(s, val(p.prev, p.next));
        return;
      }
      case 'addSheet':
      case 'removeSheet': {
        const adding = (p.k === 'addSheet') === (dir === 'forward');
        const ids = [...this.meta.sheetIds];
        if (adding) {
          this.sheets.set(p.sheet.id, p.sheet);
          ids.splice(p.index, 0, p.sheet.id);
        } else {
          this.sheets.delete(p.sheet.id);
          const i = ids.indexOf(p.sheet.id);
          if (i >= 0) ids.splice(i, 1);
          if (this.meta.activeSheet === p.sheet.id) this.meta = { ...this.meta, activeSheet: ids[0] };
        }
        this.meta = { ...this.meta, sheetIds: ids };
        return;
      }
      case 'wb':
        this.meta = { ...this.meta, ...val(p.prev, p.next) };
        return;
    }
  }

  private afterChange(patches: Patch[], _dir: 'forward' | 'backward') {
    const info: CommitInfo = { rows: [], sheets: new Set(), removedSheets: [], wb: false, structural: false };
    const seenRows = new Set<string>();
    for (const p of patches) {
      switch (p.k) {
        case 'row': {
          const key = p.sheetId + '/' + p.rowId;
          if (seenRows.has(key)) break;
          seenRows.add(key);
          info.rows.push({ sheetId: p.sheetId, rowId: p.rowId, row: this.sheets.get(p.sheetId)?.rows.get(p.rowId) });
          if (!p.prev || !p.next || p.prev.h !== p.next.h || p.prev.hidden !== p.next.hidden) info.structural = true;
          break;
        }
        case 'order':
        case 'cols':
        case 'sheet':
          info.sheets.add(p.sheetId);
          info.structural = true;
          break;
        case 'addSheet':
        case 'removeSheet':
          info.wb = true;
          info.structural = true;
          info.sheets.add(p.sheet.id);
          if (!this.sheets.has(p.sheet.id)) info.removedSheets.push(p.sheet.id);
          else for (const [rowId, row] of p.sheet.rows) info.rows.push({ sheetId: p.sheet.id, rowId, row });
          break;
        case 'wb':
          info.wb = true;
          break;
      }
    }
    this.invalidateCalc();
    if (info.structural) {
      this.layoutVersion++;
      this.views.clear();
    }
    for (const l of this.commitListeners) l(info);
    this.emit();
  }

  invalidateCalc() {
    this.calcEpoch++;
    this.calcCache.clear();
    this.lookupCache.clear();
  }

  // ─── низкоуровневые изменения ──────────────────────────────────────────────

  private putRow(sheet: Sheet, row: Row) {
    const prev = sheet.rows.get(row.id);
    sheet.rows.set(row.id, row);
    this.record({ k: 'row', sheetId: sheet.id, rowId: row.id, prev, next: row });
  }

  private dropRow(sheet: Sheet, rowId: RowId) {
    const prev = sheet.rows.get(rowId);
    sheet.rows.delete(rowId);
    this.record({ k: 'row', sheetId: sheet.id, rowId, prev, next: undefined });
  }

  private setOrder(sheet: Sheet, next: RowId[]) {
    const prev = sheet.rowOrder;
    sheet.rowOrder = next;
    this.record({ k: 'order', sheetId: sheet.id, prev, next });
  }

  private setColumns(sheet: Sheet, next: Column[]) {
    const prev = sheet.columns;
    sheet.columns = next;
    this.record({ k: 'cols', sheetId: sheet.id, prev, next });
  }

  setSheetMeta(sheet: Sheet, patch: Partial<SheetMeta>) {
    const prev: Partial<SheetMeta> = {};
    for (const k of Object.keys(patch) as (keyof SheetMeta)[]) (prev as Record<string, unknown>)[k] = sheet[k];
    Object.assign(sheet, patch);
    this.record({ k: 'sheet', sheetId: sheet.id, prev, next: patch });
  }

  setWorkbookMeta(patch: Partial<WorkbookMeta>) {
    const prev: Partial<WorkbookMeta> = {};
    for (const k of Object.keys(patch) as (keyof WorkbookMeta)[]) (prev as Record<string, unknown>)[k] = this.meta[k];
    this.meta = { ...this.meta, ...patch };
    this.record({ k: 'wb', prev, next: patch });
  }

  /** Записать ячейку (undefined или пустой объект — очистить). */
  setCell(sheet: Sheet, rowId: RowId, colId: ColId, cell: Cell | undefined) {
    const row = sheet.rows.get(rowId);
    if (!row) return;
    const cells = { ...row.cells };
    if (!cell || isEmptyCell(cell)) delete cells[colId];
    else cells[colId] = cell;
    this.putRow(sheet, { ...row, cells });
  }

  patchCell(sheet: Sheet, rowId: RowId, colId: ColId, fn: (c: Cell) => Cell | undefined) {
    const row = sheet.rows.get(rowId);
    if (!row) return;
    this.setCell(sheet, rowId, colId, fn({ ...(row.cells[colId] ?? {}) }));
  }

  patchRow(sheet: Sheet, rowId: RowId, patch: Partial<Row>) {
    const row = sheet.rows.get(rowId);
    if (!row) return;
    this.putRow(sheet, { ...row, ...patch });
  }

  // ─── вычисления ────────────────────────────────────────────────────────────

  compile(src: string): Node | Error {
    let c = this.compiled.get(src);
    if (!c) {
      try {
        c = parseFormula(src);
      } catch (e) {
        c = e instanceof Error ? e : new Error(String(e));
      }
      if (this.compiled.size > 60000) this.compiled.clear();
      this.compiled.set(src, c);
    }
    return c;
  }

  sheetByName(name: string): string | null {
    const low = name.toLocaleLowerCase('ru');
    for (const s of this.sheets.values()) if (s.name.toLocaleLowerCase('ru') === low) return s.id;
    return null;
  }

  rowCount(sheetId: string): number {
    return this.sheets.get(sheetId)?.rowOrder.length ?? 0;
  }

  colCount(sheetId: string): number {
    return this.sheets.get(sheetId)?.columns.length ?? 0;
  }

  namedValue(name: string): Scalar | undefined {
    const low = name.toLocaleLowerCase('ru');
    const nv = this.meta.names.find((n) => n.name.toLocaleLowerCase('ru') === low);
    return nv ? nv.value : undefined;
  }

  rowHidden(sheetId: string, row: number): { manual: boolean; filtered: boolean } {
    const s = this.sheets.get(sheetId);
    const id = s?.rowOrder[row];
    if (!s || !id) return { manual: false, filtered: false };
    const manual = !!s.rows.get(id)?.hidden;
    const view = this.view(s);
    return { manual, filtered: !manual && view.filtered && !view.visible.has(id) };
  }

  lookupIndex(range: RangeVal, col: number): Map<string, number> {
    const key = `${range.sheetId}:${range.top}:${range.left + col}:${range.rows}`;
    let m = this.lookupCache.get(key);
    if (!m) {
      m = buildLookupIndex(range, col);
      this.lookupCache.set(key, m);
    }
    return m;
  }

  /** Вычисленное значение ячейки по физическим координатам. */
  cellValue(sheetId: string, r: number, c: number): Scalar {
    const sheet = this.sheets.get(sheetId);
    if (!sheet) return ERR.REF;
    const rowId = sheet.rowOrder[r];
    const col = sheet.columns[c];
    if (!rowId || !col) return null;
    const cell = sheet.rows.get(rowId)?.cells[col.id];
    if (cell?.f) return this.evalFormula(sheet, r, c, cell.f, 0);
    if (cell && cell.v !== undefined) return cell.v;
    if (col.formula && !cell?.img) {
      // в совсем пустой строке формулы столбца не считаем — иначе новые строки полны нулей
      const cells = sheet.rows.get(rowId)?.cells;
      let filled = false;
      if (cells) for (const k in cells) if (cells[k].v !== undefined || cells[k].f || cells[k].img) {
        filled = true;
        break;
      }
      return filled ? this.evalFormula(sheet, r, c, col.formula, r) : null;
    }
    return null;
  }

  private evalFormula(sheet: Sheet, r: number, c: number, src: string, baseRow: number): Scalar {
    const key = `${sheet.id}${r}${c}`;
    const cached = this.calcCache.get(key);
    if (cached !== undefined) return cached;
    if (this.calcStack.has(key)) return ERR.CYCLE;
    const node = this.compile(src);
    if (node instanceof Error) {
      this.calcCache.set(key, ERR.NAME);
      return ERR.NAME;
    }
    this.calcStack.add(key);
    let v: Scalar;
    try {
      const ctx: EvalCtx = { host: this, sheetId: sheet.id, row: r, col: c, baseRow };
      v = evaluateToScalar(node, ctx);
      if (v === null) v = 0;
    } catch (e) {
      console.error(e);
      v = ERR.VALUE;
    } finally {
      this.calcStack.delete(key);
    }
    this.calcCache.set(key, v);
    return v;
  }

  /** Формула, которую видит пользователь в ячейке: своя или формула столбца, сдвинутая на строку. */
  formulaAt(sheet: Sheet, r: number, c: number): string | null {
    const rowId = sheet.rowOrder[r];
    const col = sheet.columns[c];
    if (!rowId || !col) return null;
    const cell = sheet.rows.get(rowId)?.cells[col.id];
    if (cell?.f) return cell.f;
    if (cell && (cell.v !== undefined || cell.img)) return null;
    if (col.formula) return shiftFormula(col.formula, r, 0);
    return null;
  }

  styleOf(_sheet: Sheet, col: Column, cell: Cell | undefined): CellStyle {
    if (!col.st) return cell?.st ?? HIDDEN_STYLE;
    if (!cell?.st) return col.st;
    return { ...col.st, ...cell.st };
  }

  display(sheet: Sheet, r: number, c: number): CellDisplay {
    const rowId = sheet.rowOrder[r];
    const col = sheet.columns[c];
    const cell = rowId && col ? sheet.rows.get(rowId)?.cells[col.id] : undefined;
    const style = col ? this.styleOf(sheet, col, cell) : HIDDEN_STYLE;
    const value = this.cellValue(sheet.id, r, c);
    const isFormula = !!cell?.f || (!!col?.formula && !(cell && (cell.v !== undefined || cell.img)));
    const text = formatScalar(value, style.nf);
    let href = cell?.href;
    if (!href && typeof value === 'string' && isUrl(value)) href = value;
    return {
      text,
      value,
      style,
      img: cell?.img,
      href,
      isNumber: typeof value === 'number',
      isError: isErr(value),
      isFormula,
      note: cell?.note,
    };
  }

  // ─── представление: фильтры, поиск, скрытые строки ─────────────────────────

  view(sheet: Sheet): SheetView {
    const key = `${this.layoutVersion}|${this.filterEpoch}|${this.search}|${sheet.rowOrder.length}|${sheet.columns.length}`;
    const cached = this.views.get(sheet.id);
    if (cached && cached.key === key) return cached.view;
    const view = this.computeView(sheet);
    this.views.set(sheet.id, { key, view });
    return view;
  }

  private computeView(sheet: Sheet): SheetView {
    const cols: number[] = [];
    sheet.columns.forEach((c, i) => {
      if (!c.hidden) cols.push(i);
    });
    const rowIndex = new Map<RowId, number>();
    sheet.rowOrder.forEach((id, i) => rowIndex.set(id, i));

    const filterEntries = Object.entries(sheet.filters).filter(([colId]) => this.colIndexOf(sheet, colId) >= 0);
    const tests = filterEntries.map(([colId, spec]) => this.makeRowTest(sheet, this.colIndexOf(sheet, colId), spec));
    const q = this.search.trim().toLocaleLowerCase('ru');
    const pinned = this.pinned.get(sheet.id);
    const filtered = tests.length > 0 || q !== '';

    const rows: RowId[] = [];
    for (let i = 0; i < sheet.rowOrder.length; i++) {
      const id = sheet.rowOrder[i];
      const row = sheet.rows.get(id);
      if (row?.hidden) continue;
      if (filtered && !pinned?.has(id)) {
        let ok = true;
        for (const t of tests) if (!t(i, row)) {
          ok = false;
          break;
        }
        if (ok && q) ok = this.rowMatchesSearch(sheet, i, cols, q);
        if (!ok) continue;
      }
      rows.push(id);
    }
    return { rows, rowIndex, cols, visible: new Set(rows), filtered, totalRows: sheet.rowOrder.length };
  }

  private rowMatchesSearch(sheet: Sheet, r: number, cols: number[], q: string): boolean {
    for (const c of cols) {
      const d = this.display(sheet, r, c);
      if (d.text && d.text.toLocaleLowerCase('ru').includes(q)) return true;
    }
    return false;
  }

  /** Ключ значения для фильтра «по значениям»: то, что видно в ячейке. */
  filterKey(sheet: Sheet, r: number, c: number): string {
    return this.display(sheet, r, c).text;
  }

  private makeRowTest(sheet: Sheet, c: number, spec: FilterSpec): (r: number, row: Row | undefined) => boolean {
    const col = sheet.columns[c];
    const allowed = spec.values ? new Set(spec.values) : null;
    const cond = spec.cond;
    let topThreshold = -Infinity;
    if (cond?.op === 'top') {
      const n = Math.max(1, Math.floor(Number(cond.a) || 10));
      const nums: number[] = [];
      for (let r = 0; r < sheet.rowOrder.length; r++) {
        const v = this.cellValue(sheet.id, r, c);
        if (typeof v === 'number') nums.push(v);
      }
      nums.sort((a, b) => b - a);
      topThreshold = nums.length ? nums[Math.min(n, nums.length) - 1] : Infinity;
    }
    const a = cond?.a ?? '';
    const b = cond?.b ?? '';
    const aNum = parseLooseNumber(a);
    const bNum = parseLooseNumber(b);
    const aLow = a.toLocaleLowerCase('ru');

    return (r, row) => {
      const cell = row?.cells[col.id];
      if (spec.photo) {
        const has = !!cell?.img;
        if (spec.photo === 'with' ? !has : has) return false;
      }
      if (spec.color) {
        const st = this.styleOf(sheet, col, cell);
        const actual = (spec.color.kind === 'bg' ? st.bg : st.fg) ?? null;
        if ((actual ?? null) !== spec.color.color) return false;
      }
      if (allowed) {
        const key = this.filterKey(sheet, r, c);
        if (!allowed.has(key)) return false;
      }
      if (cond) {
        const v = this.cellValue(sheet.id, r, c);
        const text = formatScalar(v, this.styleOf(sheet, col, cell).nf).toLocaleLowerCase('ru');
        const n = typeof v === 'number' ? v : typeof v === 'string' ? parseLooseNumber(v) : null;
        switch (cond.op) {
          case 'contains':
            return text.includes(aLow);
          case 'notContains':
            return !text.includes(aLow);
          case 'begins':
            return text.startsWith(aLow);
          case 'ends':
            return text.endsWith(aLow);
          case 'eq':
            return aNum !== null && n !== null ? n === aNum : text === aLow;
          case 'neq':
            return aNum !== null && n !== null ? n !== aNum : text !== aLow;
          case 'gt':
            return n !== null && aNum !== null && n > aNum;
          case 'gte':
            return n !== null && aNum !== null && n >= aNum;
          case 'lt':
            return n !== null && aNum !== null && n < aNum;
          case 'lte':
            return n !== null && aNum !== null && n <= aNum;
          case 'between':
            return n !== null && aNum !== null && bNum !== null && n >= Math.min(aNum, bNum) && n <= Math.max(aNum, bNum);
          case 'empty':
            return text === '' && !cell?.img;
          case 'notEmpty':
            return text !== '' || !!cell?.img;
          case 'top':
            return n !== null && n >= topThreshold;
        }
      }
      return true;
    };
  }

  setSearch(q: string) {
    if (q === this.search) return;
    this.search = q;
    this.filterEpoch++;
    this.pinned.clear();
    this.views.clear();
    this.emit();
  }

  setFilter(sheet: Sheet, colId: ColId, spec: FilterSpec | null) {
    const filters = { ...sheet.filters };
    if (spec && (spec.values || spec.cond || spec.color || spec.photo)) filters[colId] = spec;
    else delete filters[colId];
    this.allow(() => this.transact(spec ? 'Фильтр' : 'Сброс фильтра', () => this.setSheetMeta(sheet, { filters })));
    this.filterEpoch++;
    this.pinned.delete(sheet.id);
    this.views.clear();
    this.emit();
  }

  clearFilters(sheet: Sheet) {
    if (!Object.keys(sheet.filters).length) return;
    this.allow(() => this.transact('Сброс фильтров', () => this.setSheetMeta(sheet, { filters: {} })));
    this.filterEpoch++;
    this.pinned.delete(sheet.id);
    this.views.clear();
    this.emit();
  }

  /** Строки, добавленные при включённом фильтре, остаются видимыми до смены фильтра. */
  private pin(sheet: Sheet, ids: RowId[]) {
    let set = this.pinned.get(sheet.id);
    if (!set) this.pinned.set(sheet.id, (set = new Set()));
    for (const id of ids) set.add(id);
  }

  // ─── операции со строками ──────────────────────────────────────────────────

  private forEachFormula(fn: (sheet: Sheet, src: string, isColumnFormula: boolean) => string, only?: SheetId) {
    for (const sheet of this.sheets.values()) {
      if (only && sheet.id !== only) continue;
      for (const id of sheet.rowOrder) {
        const row = sheet.rows.get(id);
        if (!row) continue;
        let changed: Record<ColId, Cell> | null = null;
        for (const colId in row.cells) {
          const cell = row.cells[colId];
          if (!cell.f) continue;
          const next = fn(sheet, cell.f, false);
          if (next !== cell.f) {
            changed ??= { ...row.cells };
            changed[colId] = { ...cell, f: next };
          }
        }
        if (changed) this.putRow(sheet, { ...row, cells: changed });
      }
      let colsChanged = false;
      const cols = sheet.columns.map((c) => {
        if (!c.formula) return c;
        const next = fn(sheet, c.formula, true);
        if (next === c.formula) return c;
        colsChanged = true;
        return { ...c, formula: next };
      });
      if (colsChanged) this.setColumns(sheet, cols);
    }
  }

  /** Поправить ссылки во всех формулах книги после вставки/удаления строк или столбцов листа. */
  private adjustAll(target: Sheet, ch: StructuralChange) {
    const low = target.name.toLocaleLowerCase('ru');
    this.forEachFormula((sheet, src, isCol) =>
      adjustFormula(src, ch, (name) => (name ? name.toLocaleLowerCase('ru') === low : sheet.id === target.id), isCol),
    );
  }

  insertRows(sheet: Sheet, at: number, count: number, template?: (i: number) => Partial<Row>): RowId[] {
    const ids: RowId[] = [];
    this.transact(count > 1 ? `Вставка ${count} строк` : 'Вставка строки', () => {
      this.adjustAll(sheet, { axis: 'row', at, count });
      for (let i = 0; i < count; i++) {
        const id = uid(8);
        ids.push(id);
        this.putRow(sheet, { id, cells: {}, ...template?.(i) });
      }
      const order = [...sheet.rowOrder];
      order.splice(at, 0, ...ids);
      this.setOrder(sheet, order);
    });
    this.pin(sheet, ids);
    this.views.clear();
    this.emit();
    return ids;
  }

  deleteRows(sheet: Sheet, rowIds: RowId[]) {
    if (!rowIds.length) return;
    const doomed = new Set(rowIds);
    this.transact(rowIds.length > 1 ? `Удаление ${rowIds.length} строк` : 'Удаление строки', () => {
      // удаляем блоками подряд идущих позиций снизу вверх, чтобы корректно поправить ссылки
      const positions = sheet.rowOrder.map((id, i) => (doomed.has(id) ? i : -1)).filter((i) => i >= 0);
      const runs: [number, number][] = [];
      for (const p of positions) {
        const last = runs[runs.length - 1];
        if (last && last[0] + last[1] === p) last[1]++;
        else runs.push([p, 1]);
      }
      for (let k = runs.length - 1; k >= 0; k--) this.adjustAll(sheet, { axis: 'row', at: runs[k][0], count: -runs[k][1] });
      this.shrinkMerges(sheet, 'row', doomed);
      for (const id of rowIds) this.dropRow(sheet, id);
      this.setOrder(
        sheet,
        sheet.rowOrder.filter((id) => !doomed.has(id)),
      );
    });
  }

  setRowsHidden(sheet: Sheet, rowIds: RowId[], hidden: boolean) {
    this.transact(hidden ? 'Скрыть строки' : 'Показать строки', () => {
      for (const id of rowIds) {
        const row = sheet.rows.get(id);
        if (row && !!row.hidden !== hidden) this.putRow(sheet, { ...row, hidden: hidden || undefined });
      }
    });
  }

  showAllRows(sheet: Sheet) {
    const ids = sheet.rowOrder.filter((id) => sheet.rows.get(id)?.hidden);
    if (ids.length) this.setRowsHidden(sheet, ids, false);
  }

  setRowsHeight(sheet: Sheet, rowIds: RowId[], h: number | undefined) {
    this.transact('Высота строки', () => {
      for (const id of rowIds) {
        const row = sheet.rows.get(id);
        if (row) this.putRow(sheet, { ...row, h: h === undefined ? undefined : Math.max(20, Math.min(600, Math.round(h))) });
      }
    });
  }

  /**
   * Сортировка всех строк листа по столбцу (как «Сортировка» в автофильтре Excel).
   * Объединения по вертикали при сортировке снимаются (Excel в таком случае сортировать отказывается).
   * Возвращает, сколько объединений снято.
   */
  sortRows(sheet: Sheet, c: number, dir: 'asc' | 'desc', by: { kind: 'value' } | { kind: 'bg' | 'fg'; color: string | null } = { kind: 'value' }): number {
    const col = sheet.columns[c];
    if (!col) return 0;
    const vertical = (sheet.merges ?? []).filter((m) => m.r0 !== m.r1);
    const n = sheet.rowOrder.length;
    const keys: { i: number; v: Scalar; empty: boolean; colorHit: boolean }[] = [];
    for (let i = 0; i < n; i++) {
      const row = sheet.rows.get(sheet.rowOrder[i]);
      const cell = row?.cells[col.id];
      const v = this.cellValue(sheet.id, i, c);
      let colorHit = false;
      if (by.kind !== 'value') {
        const st = this.styleOf(sheet, col, cell);
        colorHit = ((by.kind === 'bg' ? st.bg : st.fg) ?? null) === by.color;
      }
      keys.push({ i, v, empty: v === null || v === '', colorHit });
    }
    const sign = dir === 'asc' ? 1 : -1;
    keys.sort((a, b) => {
      if (by.kind !== 'value') {
        if (a.colorHit !== b.colorHit) return a.colorHit ? -1 : 1;
        return a.i - b.i;
      }
      if (a.empty !== b.empty) return a.empty ? 1 : -1; // пустые всегда в конце
      if (isErr(a.v) !== isErr(b.v)) return isErr(a.v) ? 1 : -1;
      const d = compareScalars(a.v, b.v) * sign;
      return d !== 0 ? d : a.i - b.i;
    });
    const order = keys.map((k) => sheet.rowOrder[k.i]);
    this.allow(() => this.transact('Сортировка', () => {
      // формулы в переехавших строках сдвигаются вместе со строкой (как в Excel)
      keys.forEach((k, newPos) => {
        const delta = newPos - k.i;
        if (!delta) return;
        const row = sheet.rows.get(sheet.rowOrder[k.i]);
        if (!row) return;
        let cells: Record<ColId, Cell> | null = null;
        for (const colId in row.cells) {
          const cell = row.cells[colId];
          if (!cell.f) continue;
          cells ??= { ...row.cells };
          cells[colId] = { ...cell, f: shiftFormula(cell.f, delta, 0) };
        }
        if (cells) this.putRow(sheet, { ...row, cells });
      });
      this.setOrder(sheet, order);
      if (vertical.length) this.setMerges(sheet, (sheet.merges ?? []).filter((m) => m.r0 === m.r1));
    }));
    return vertical.length;
  }

  // ─── объединённые ячейки ───────────────────────────────────────────────────

  setMerges(sheet: Sheet, merges: Merge[]) {
    this.setSheetMeta(sheet, { merges: merges.length ? merges : undefined });
  }

  /** Перед удалением строк/столбцов: угол объединения, попавший под удаление, переезжает на ближайшую уцелевшую линию внутри. */
  private shrinkMerges(sheet: Sheet, axis: 'row' | 'col', doomed: Set<string>) {
    if (!sheet.merges?.length) return;
    const line = axis === 'row' ? sheet.rowOrder : sheet.columns.map((c) => c.id);
    const pos = new Map(line.map((id, i) => [id, i]));
    let changed = false;
    const next: Merge[] = [];
    for (const m of sheet.merges) {
      const [ka, kb] = axis === 'row' ? (['r0', 'r1'] as const) : (['c0', 'c1'] as const);
      const a = pos.get(m[ka]);
      const b = pos.get(m[kb]);
      if (a === undefined || b === undefined) {
        changed = true;
        continue;
      }
      if (!doomed.has(m[ka]) && !doomed.has(m[kb])) {
        next.push(m);
        continue;
      }
      changed = true;
      const lo = Math.min(a, b);
      const hi = Math.max(a, b);
      let first = -1;
      let last = -1;
      for (let i = lo; i <= hi; i++) {
        if (doomed.has(line[i])) continue;
        if (first < 0) first = i;
        last = i;
      }
      if (first < 0) continue;
      const moved = { ...m, [ka]: line[first], [kb]: line[last] };
      // от объединения осталась одна ячейка — оно больше не нужно
      if (moved.r0 === moved.r1 && moved.c0 === moved.c1) continue;
      next.push(moved);
    }
    if (changed) this.setMerges(sheet, next);
  }

  // ─── операции со столбцами ─────────────────────────────────────────────────

  insertColumns(sheet: Sheet, at: number, count = 1, names?: string[]): ColId[] {
    const ids: ColId[] = [];
    this.transact('Вставка столбца', () => {
      this.adjustAll(sheet, { axis: 'col', at, count });
      const cols = [...sheet.columns];
      const fresh: Column[] = [];
      for (let i = 0; i < count; i++) {
        const id = uid(6);
        ids.push(id);
        fresh.push({ id, name: names?.[i] ?? this.nextColumnName(sheet, i), w: DEFAULT_COL_WIDTH });
      }
      cols.splice(at, 0, ...fresh);
      this.setColumns(sheet, cols);
      if (at < sheet.frozen) this.setSheetMeta(sheet, { frozen: sheet.frozen + count });
    });
    return ids;
  }

  private nextColumnName(sheet: Sheet, offset: number): string {
    const base = 'Столбец';
    const used = new Set(sheet.columns.map((c) => c.name));
    let k = sheet.columns.length + 1 + offset;
    while (used.has(`${base} ${k}`)) k++;
    return `${base} ${k}`;
  }

  deleteColumns(sheet: Sheet, colIds: ColId[]) {
    const doomed = new Set(colIds);
    if (doomed.size >= sheet.columns.length) return;
    this.transact(colIds.length > 1 ? 'Удаление столбцов' : 'Удаление столбца', () => {
      const positions = sheet.columns.map((c, i) => (doomed.has(c.id) ? i : -1)).filter((i) => i >= 0);
      for (let k = positions.length - 1; k >= 0; k--) this.adjustAll(sheet, { axis: 'col', at: positions[k], count: -1 });
      this.shrinkMerges(sheet, 'col', doomed);
      for (const id of sheet.rowOrder) {
        const row = sheet.rows.get(id);
        if (!row) continue;
        let hit = false;
        for (const cid of colIds) if (row.cells[cid]) hit = true;
        if (!hit) continue;
        const cells = { ...row.cells };
        for (const cid of colIds) delete cells[cid];
        this.putRow(sheet, { ...row, cells });
      }
      const frozenLost = sheet.columns.slice(0, sheet.frozen).filter((c) => doomed.has(c.id)).length;
      this.setColumns(
        sheet,
        sheet.columns.filter((c) => !doomed.has(c.id)),
      );
      const filters = { ...sheet.filters };
      let fchanged = false;
      for (const cid of colIds) if (filters[cid]) {
        delete filters[cid];
        fchanged = true;
      }
      const patch: Partial<SheetMeta> = {};
      if (fchanged) patch.filters = filters;
      if (frozenLost) patch.frozen = Math.max(0, sheet.frozen - frozenLost);
      if (sheet.keyColId && doomed.has(sheet.keyColId)) patch.keyColId = sheet.columns[0]?.id;
      if (Object.keys(patch).length) this.setSheetMeta(sheet, patch);
    });
  }

  moveColumn(sheet: Sheet, from: number, to: number) {
    if (from === to) return;
    const perm = sheet.columns.map((_, i) => i);
    const [moved] = perm.splice(from, 1);
    perm.splice(to, 0, moved);
    const newIndexOf = new Map(perm.map((old, idx) => [old, idx]));
    const low = sheet.name.toLocaleLowerCase('ru');
    this.transact('Перемещение столбца', () => {
      this.forEachFormula((s, src) =>
        remapColumns(src, (c) => newIndexOf.get(c) ?? c, (name) => (name ? name.toLocaleLowerCase('ru') === low : s.id === sheet.id)),
      );
      this.setColumns(
        sheet,
        perm.map((i) => sheet.columns[i]),
      );
    });
  }

  updateColumn(sheet: Sheet, colId: ColId, patch: Partial<Column>, label = 'Столбец') {
    this.transact(label, () => {
      this.setColumns(
        sheet,
        sheet.columns.map((c) => (c.id === colId ? { ...c, ...patch } : c)),
      );
    });
  }

  setColumnsHidden(sheet: Sheet, colIds: ColId[], hidden: boolean) {
    const set = new Set(colIds);
    // хотя бы один столбец должен оставаться видимым
    if (hidden && sheet.columns.every((c) => c.hidden || set.has(c.id))) return;
    this.transact(hidden ? 'Скрыть столбцы' : 'Показать столбцы', () => {
      this.setColumns(
        sheet,
        sheet.columns.map((c) => (set.has(c.id) ? { ...c, hidden: hidden || undefined } : c)),
      );
    });
  }

  // ─── листы ─────────────────────────────────────────────────────────────────

  addSheet(sheet: Sheet, activate = true) {
    this.transact('Новый лист', () => {
      this.record({ k: 'addSheet', sheet, index: this.meta.sheetIds.length });
      this.sheets.set(sheet.id, sheet);
      this.meta = { ...this.meta, sheetIds: [...this.meta.sheetIds, sheet.id] };
      if (activate) this.setWorkbookMeta({ activeSheet: sheet.id });
    });
  }

  removeSheet(sheetId: SheetId) {
    if (this.meta.sheetIds.length <= 1) return;
    const sheet = this.sheet(sheetId);
    const index = this.meta.sheetIds.indexOf(sheetId);
    this.transact('Удаление листа', () => {
      if (this.meta.activeSheet === sheetId) {
        const next = this.meta.sheetIds[index + 1] ?? this.meta.sheetIds[index - 1];
        this.setWorkbookMeta({ activeSheet: next });
      }
      this.record({ k: 'removeSheet', sheet, index });
      this.sheets.delete(sheetId);
      this.meta = { ...this.meta, sheetIds: this.meta.sheetIds.filter((id) => id !== sheetId) };
    });
  }

  renameSheet(sheet: Sheet, name: string) {
    const clean = name.trim().slice(0, 60);
    if (!clean || clean === sheet.name) return;
    const from = sheet.name;
    this.transact('Переименование листа', () => {
      this.forEachFormula((_s, src) => renameSheetRefs(src, from, clean));
      this.setSheetMeta(sheet, { name: clean });
    });
  }

  setActiveSheet(sheetId: SheetId) {
    if (this.meta.activeSheet === sheetId) return;
    // переключение листа не попадает в историю отмены
    this.meta = { ...this.meta, activeSheet: sheetId };
    for (const l of this.commitListeners) l({ rows: [], sheets: new Set(), removedSheets: [], wb: true, structural: false });
    this.layoutVersion++;
    this.emit();
  }

  setNames(names: NamedValue[]) {
    this.transact('Именованные значения', () => this.setWorkbookMeta({ names }));
  }

  /** Заменить всю книгу (восстановление из копии или импорт). История очищается. */
  replaceAll(meta: WorkbookMeta, sheets: Sheet[]) {
    this.meta = meta;
    this.sheets = new Map(sheets.map((s) => [s.id, s]));
    this.undoStack = [];
    this.redoStack = [];
    this.pinned.clear();
    this.views.clear();
    this.filterEpoch++;
    this.invalidateCalc();
    this.layoutVersion++;
    this.emit();
  }

  /** Уведомить подписчиков без записи в историю (например, после загрузки фото). */
  touch() {
    this.emit();
  }
}

export function isEmptyCell(c: Cell): boolean {
  return (
    (c.v === undefined || c.v === '') &&
    !c.f &&
    !c.img &&
    !c.href &&
    !c.note &&
    (!c.st || Object.values(c.st).every((x) => x === undefined))
  );
}
