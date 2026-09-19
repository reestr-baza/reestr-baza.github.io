import { colToLetters } from '../formula/a1';
import { shiftFormula } from '../formula/refs';
import { editText, parseInput } from '../model/format';
import { isEmptyCell } from '../model/store';
import type { Cell, CellStyle, ColId, NumFmt, Row, RowId, Sheet } from '../model/types';
import { importImage, isImageFile } from '../storage/images';
import { selRect, useUI, type Selection } from '../ui/state';
import { gridApi } from './gridApi';
import { store } from './instance';

// ─── контекст ────────────────────────────────────────────────────────────────

export function ctx() {
  const sheet = store.activeSheet;
  const view = store.view(sheet);
  return { sheet, view };
}

export function cellAt(vr: number, vc: number) {
  const { sheet, view } = ctx();
  const rowId = view.rows[vr];
  const c = view.cols[vc];
  const col = sheet.columns[c];
  if (rowId === undefined || !col) return null;
  return { rowId, colId: col.id, c, phys: view.rowIndex.get(rowId)!, col, row: sheet.rows.get(rowId) };
}

export function clampSel(s: Selection): Selection {
  const { view } = ctx();
  const maxR = Math.max(0, view.rows.length - 1);
  const maxC = Math.max(0, view.cols.length - 1);
  const cl = (v: number, m: number) => Math.max(0, Math.min(m, v));
  return { ar: cl(s.ar, maxR), ac: cl(s.ac, maxC), fr: cl(s.fr, maxR), fc: cl(s.fc, maxC) };
}

export function setSelection(s: Selection, scroll = true) {
  const next = clampSel(s);
  useUI.getState().setSel(next);
  if (!scroll) return;
  // выделен столбец/строка целиком — держим в поле зрения активную ячейку, а не дальний угол
  const { view } = ctx();
  const { r1, r2, c1, c2 } = selRect(next);
  const wholeCols = r1 === 0 && r2 === view.rows.length - 1 && r2 > 0;
  const wholeRows = c1 === 0 && c2 === view.cols.length - 1 && c2 > 0;
  gridApi.scrollToCell(wholeCols ? next.ar : next.fr, wholeRows ? next.ac : next.fc);
}

export function selectAll() {
  const { view } = ctx();
  setSelection({ ar: 0, ac: 0, fr: view.rows.length - 1, fc: view.cols.length - 1 }, false);
}

/** Все видимые ячейки выделения. */
export function forEachSelected(fn: (x: NonNullable<ReturnType<typeof cellAt>>, vr: number, vc: number) => void) {
  const { r1, r2, c1, c2 } = selRect(useUI.getState().sel);
  for (let vr = r1; vr <= r2; vr++)
    for (let vc = c1; vc <= c2; vc++) {
      const x = cellAt(vr, vc);
      if (x) fn(x, vr, vc);
    }
}

export function selectionCoversAllRows(): boolean {
  const { view } = ctx();
  const { r1, r2 } = selRect(useUI.getState().sel);
  return r1 === 0 && r2 === view.rows.length - 1 && !view.filtered && view.rows.length === view.totalRows && view.rows.length > 1;
}

export function selectedRowIds(): RowId[] {
  const { view } = ctx();
  const { r1, r2 } = selRect(useUI.getState().sel);
  return view.rows.slice(r1, r2 + 1);
}

export function selectedColIds(): ColId[] {
  const { sheet, view } = ctx();
  const { c1, c2 } = selRect(useUI.getState().sel);
  return view.cols.slice(c1, c2 + 1).map((c) => sheet.columns[c].id);
}

export function toast(text: string, extra?: { tone?: 'error' | 'plain'; action?: { label: string; run: () => void } }) {
  useUI.getState().toast({ text, ...extra });
}

// ─── навигация ───────────────────────────────────────────────────────────────

export function moveSel(dr: number, dc: number, extend = false) {
  const s = useUI.getState().sel;
  if (extend) setSelection({ ...s, fr: s.fr + dr, fc: s.fc + dc });
  else setSelection({ ar: s.ar + dr, ac: s.ac + dc, fr: s.ar + dr, fc: s.ac + dc });
}

function isBlankAt(vr: number, vc: number): boolean {
  const x = cellAt(vr, vc);
  if (!x) return true;
  const d = store.display(ctx().sheet, x.phys, x.c);
  return d.text === '' && !d.img;
}

/** Ctrl+стрелка: к краю блока данных, как в Excel. */
export function jump(dr: number, dc: number, extend = false) {
  const { view } = ctx();
  const s = useUI.getState().sel;
  let r = extend ? s.fr : s.ar;
  let c = extend ? s.fc : s.ac;
  const maxR = view.rows.length - 1;
  const maxC = view.cols.length - 1;
  const inside = (rr: number, cc: number) => rr >= 0 && rr <= maxR && cc >= 0 && cc <= maxC;
  const nextBlank = !inside(r + dr, c + dc) || isBlankAt(r + dr, c + dc);
  if (isBlankAt(r, c) || nextBlank) {
    // ищем первую непустую
    r += dr;
    c += dc;
    while (inside(r, c) && isBlankAt(r, c)) {
      r += dr;
      c += dc;
    }
    if (!inside(r, c)) {
      r = Math.max(0, Math.min(maxR, r));
      c = Math.max(0, Math.min(maxC, c));
    }
  } else {
    while (inside(r + dr, c + dc) && !isBlankAt(r + dr, c + dc)) {
      r += dr;
      c += dc;
    }
  }
  if (extend) setSelection({ ...s, fr: r, fc: c });
  else setSelection({ ar: r, ac: c, fr: r, fc: c });
}

// ─── редактирование ──────────────────────────────────────────────────────────

/** Текст, который увидит пользователь при редактировании ячейки. */
export function editableText(vr: number, vc: number): string {
  const x = cellAt(vr, vc);
  if (!x) return '';
  const { sheet } = ctx();
  const f = store.formulaAt(sheet, x.phys, x.c);
  if (f) return f;
  const cell = x.row?.cells[x.colId];
  const st = store.styleOf(sheet, x.col, cell);
  return editText(cell?.v, st.nf);
}

export function startEdit(mode: 'enter' | 'edit', initial?: string, source: 'cell' | 'bar' = 'cell') {
  const { sel } = useUI.getState();
  const x = cellAt(sel.ar, sel.ac);
  if (!x) return;
  const text = initial ?? editableText(sel.ar, sel.ac);
  useUI.getState().set({ edit: { r: sel.ar, c: sel.ac, text, mode, source }, menu: null });
}

export function cancelEdit() {
  useUI.getState().set({ edit: null });
  gridApi.focus();
}

/** Записать введённый текст в ячейку. Возвращает false, если в формуле ошибка. */
export function writeInput(vr: number, vc: number, text: string, label = 'Ввод'): boolean {
  const x = cellAt(vr, vc);
  if (!x) return true;
  const { sheet } = ctx();
  const parsed = parseInput(text);
  if (parsed.f) {
    const node = store.compile(parsed.f);
    if (node instanceof Error) {
      toast(`В формуле ошибка: ${node.message}`, { tone: 'error' });
      return false;
    }
  }
  const prevCell = x.row?.cells[x.colId];
  const colFormulaShown = !prevCell?.f && x.col.formula ? shiftFormula(x.col.formula, x.phys, 0) : null;
  // ввод той же формулы, что уже задана для столбца, не создаёт исключение
  if (parsed.f && colFormulaShown && parsed.f === colFormulaShown && !prevCell?.v) return true;

  store.transact(label, () => {
    store.patchCell(sheet, x.rowId, x.colId, (c) => {
      const next: Cell = { ...c };
      delete next.v;
      delete next.f;
      if (parsed.f) next.f = parsed.f;
      else if (parsed.v !== undefined) next.v = parsed.v;
      if (parsed.nf) {
        const cur = store.styleOf(sheet, x.col, c).nf;
        if (!cur || cur.k === 'general') next.st = { ...next.st, nf: parsed.nf };
      }
      // пустой ввод в ячейке столбца с формулой — вернуть формулу столбца
      return next;
    });
  });
  if (parsed.f && !x.col.formula) offerColumnFormula(vr, vc, parsed.f);
  return true;
}

let lastOffer = '';
function offerColumnFormula(vr: number, vc: number, f: string) {
  const x = cellAt(vr, vc);
  if (!x) return;
  const { sheet } = ctx();
  if (sheet.rowOrder.length < 3) return;
  // формула с относительными ссылками на ту же строку — кандидат на формулу столбца
  const anchored = shiftFormula(f, -x.phys, 0);
  if (anchored.includes('#ССЫЛКА!') || lastOffer === x.colId + anchored) return;
  lastOffer = x.colId + anchored;
  toast(`Формула записана в ${colToLetters(x.c)}${x.phys + 1}`, {
    action: { label: 'Применить ко всему столбцу', run: () => applyColumnFormula(x.colId, anchored) },
  });
}

export function applyColumnFormula(colId: ColId, anchored: string | undefined) {
  const { sheet } = ctx();
  const c = store.colIndexOf(sheet, colId);
  if (c < 0) return;
  store.transact(anchored ? 'Формула столбца' : 'Убрать формулу столбца', () => {
    store.updateColumn(sheet, colId, { formula: anchored });
    if (!anchored) return;
    // ячейки с той же формулой больше не нужны — их заменит формула столбца
    sheet.rowOrder.forEach((rowId, i) => {
      const cell = sheet.rows.get(rowId)?.cells[colId];
      if (cell?.f && shiftFormula(cell.f, -i, 0) === anchored) {
        store.patchCell(sheet, rowId, colId, (cc) => {
          delete cc.f;
          return cc;
        });
      }
    });
  });
  toast(anchored ? 'Формула применена ко всем строкам столбца' : 'Формула столбца убрана');
}

export function commitEdit(move: 'down' | 'up' | 'right' | 'left' | null = 'down'): boolean {
  const { edit } = useUI.getState();
  if (!edit) return true;
  const ok = writeInput(edit.r, edit.c, edit.text);
  if (!ok) return false;
  useUI.getState().set({ edit: null });
  const s = useUI.getState().sel;
  const { r1, r2, c1, c2 } = selRect(s);
  const multi = r1 !== r2 || c1 !== c2;
  if (move) {
    const dr = move === 'down' ? 1 : move === 'up' ? -1 : 0;
    const dc = move === 'right' ? 1 : move === 'left' ? -1 : 0;
    if (multi) {
      // ввод внутри выделенного диапазона: курсор ходит по нему, как в Excel
      let r = s.ar + dr;
      let c = s.ac + dc;
      if (r > r2) {
        r = r1;
        c = c + 1 > c2 ? c1 : c + 1;
      }
      if (c > c2) {
        c = c1;
        r = r + 1 > r2 ? r1 : r + 1;
      }
      if (r < r1) r = r2;
      if (c < c1) c = c2;
      useUI.getState().setSel({ ...s, ar: r, ac: c });
    } else moveSel(dr, dc);
  }
  gridApi.focus();
  return true;
}

// ─── оформление ──────────────────────────────────────────────────────────────

type StylePatch = Partial<Record<keyof CellStyle, CellStyle[keyof CellStyle] | undefined>>;

function mergeStyle(st: CellStyle | undefined, patch: StylePatch): CellStyle | undefined {
  const next: CellStyle = { ...st };
  for (const [k, v] of Object.entries(patch)) {
    if (v === undefined || v === false) delete (next as Record<string, unknown>)[k];
    else (next as Record<string, unknown>)[k] = v;
  }
  return Object.keys(next).length ? next : undefined;
}

/** Применить оформление к выделению. Выделены столбцы целиком — пишем в оформление столбца. */
export function applyStyle(patch: StylePatch, label = 'Оформление') {
  const { sheet } = ctx();
  if (selectionCoversAllRows()) {
    const colIds = new Set(selectedColIds());
    store.transact(label, () => {
      for (const colId of colIds) {
        const col = sheet.columns.find((c) => c.id === colId)!;
        store.updateColumn(sheet, colId, { st: mergeStyle(col.st, patch) });
        // свои значения ячеек по тем же свойствам перекрыли бы столбец — снимаем их
        for (const rowId of sheet.rowOrder) {
          const cell = sheet.rows.get(rowId)?.cells[colId];
          if (!cell?.st) continue;
          if (!Object.keys(patch).some((k) => k in cell.st!)) continue;
          const strip: StylePatch = {};
          for (const k of Object.keys(patch)) strip[k as keyof CellStyle] = undefined;
          store.patchCell(sheet, rowId, colId, (c) => ({ ...c, st: mergeStyle(c.st, strip) }));
        }
      }
    });
    return;
  }
  store.transact(label, () => {
    forEachSelected((x) => {
      store.patchCell(sheet, x.rowId, x.colId, (c) => ({ ...c, st: mergeStyle(c.st, patch) }));
    });
  });
}

/** Текущее оформление активной ячейки (с учётом столбца). */
export function activeStyle(): CellStyle {
  const { sel } = useUI.getState();
  const x = cellAt(sel.ar, sel.ac);
  if (!x) return {};
  return store.styleOf(ctx().sheet, x.col, x.row?.cells[x.colId]);
}

export function toggleStyle(key: 'b' | 'i' | 'u' | 's') {
  const on = !activeStyle()[key];
  const labels = { b: 'Жирный', i: 'Курсив', u: 'Подчёркнутый', s: 'Зачёркнутый' };
  applyStyle({ [key]: on || undefined }, labels[key]);
}

export function setNumFmt(nf: NumFmt | undefined) {
  applyStyle({ nf: nf && nf.k !== 'general' ? nf : undefined }, 'Формат чисел');
}

export function clearSelection(what: 'contents' | 'formats' | 'all' = 'contents') {
  const { sheet } = ctx();
  const labels = { contents: 'Очистка', formats: 'Очистка формата', all: 'Очистка всего' };
  store.transact(labels[what], () => {
    forEachSelected((x) => {
      store.patchCell(sheet, x.rowId, x.colId, (c) => {
        if (what === 'formats') return { ...c, st: undefined };
        if (what === 'all') return undefined;
        const next = { ...c };
        delete next.v;
        delete next.f;
        delete next.img;
        delete next.href;
        return next;
      });
    });
  });
}

// ─── строки и столбцы ────────────────────────────────────────────────────────

export function insertRows(where: 'above' | 'below', count?: number) {
  const { sheet, view } = ctx();
  const { r1, r2 } = selRect(useUI.getState().sel);
  const n = count ?? r2 - r1 + 1;
  const anchorId = view.rows[where === 'above' ? r1 : r2];
  const at = anchorId === undefined ? sheet.rowOrder.length : view.rowIndex.get(anchorId)! + (where === 'below' ? 1 : 0);
  store.insertRows(sheet, at, n);
  const s = useUI.getState().sel;
  if (where === 'above') setSelection({ ...s, ar: r1, fr: r1 + n - 1 });
  else setSelection({ ...s, ar: r2 + 1, fr: r2 + n });
}

export function appendRow(): number {
  const { sheet } = ctx();
  const ids = store.insertRows(sheet, sheet.rowOrder.length, 1);
  const v = store.view(sheet);
  return v.rows.indexOf(ids[0]);
}

export function deleteSelectedRows() {
  const { sheet, view } = ctx();
  const ids = selectedRowIds();
  if (!ids.length) return;
  if (ids.length === view.rows.length && view.rows.length === sheet.rowOrder.length) {
    // последнюю строку листа не удаляем — оставляем пустую
    store.transact('Удаление строк', () => {
      store.insertRows(sheet, sheet.rowOrder.length, 1);
      store.deleteRows(sheet, ids);
    });
  } else store.deleteRows(sheet, ids);
  const s = useUI.getState().sel;
  const r = Math.min(s.ar, s.fr);
  setSelection({ ar: r, ac: s.ac, fr: r, fc: s.ac });
  toast(ids.length > 1 ? `Удалено строк: ${ids.length}` : 'Строка удалена', { action: { label: 'Вернуть', run: undo } });
}

export function insertColumns(where: 'left' | 'right') {
  const { sheet, view } = ctx();
  const { c1, c2 } = selRect(useUI.getState().sel);
  const n = c2 - c1 + 1;
  if (sheet.columns.length + n > 200) {
    toast('В листе может быть не больше 200 столбцов', { tone: 'error' });
    return;
  }
  const at = where === 'left' ? view.cols[c1] : view.cols[c2] + 1;
  store.insertColumns(sheet, at, n);
  const s = useUI.getState().sel;
  const vc = where === 'left' ? c1 : c2 + 1;
  setSelection({ ...s, ac: vc, fc: vc + n - 1 });
}

export function deleteSelectedColumns() {
  const { sheet } = ctx();
  const ids = selectedColIds();
  if (ids.length >= sheet.columns.length) {
    toast('Нельзя удалить все столбцы листа', { tone: 'error' });
    return;
  }
  store.deleteColumns(sheet, ids);
  toast(ids.length > 1 ? `Удалено столбцов: ${ids.length}` : 'Столбец удалён', { action: { label: 'Вернуть', run: undo } });
  setSelection(useUI.getState().sel);
}

export function hideSelectedRows() {
  store.setRowsHidden(ctx().sheet, selectedRowIds(), true);
  setSelection(useUI.getState().sel);
}

export function hideSelectedColumns() {
  const { sheet } = ctx();
  const ids = selectedColIds();
  if (sheet.columns.filter((c) => !c.hidden).length <= ids.length) {
    toast('Хотя бы один столбец должен остаться видимым', { tone: 'error' });
    return;
  }
  store.setColumnsHidden(sheet, ids, true);
  setSelection(useUI.getState().sel);
}

/** Показать скрытые строки внутри выделения (или все, если выделена одна строка). */
export function unhideRowsAround() {
  const { sheet, view } = ctx();
  const { r1, r2 } = selRect(useUI.getState().sel);
  const from = r1 > 0 ? view.rowIndex.get(view.rows[r1 - 1])! : 0;
  const to = r2 + 1 < view.rows.length ? view.rowIndex.get(view.rows[r2 + 1])! : sheet.rowOrder.length - 1;
  const ids = sheet.rowOrder.slice(from, to + 1).filter((id) => sheet.rows.get(id)?.hidden);
  if (ids.length) store.setRowsHidden(sheet, ids, false);
}

export function unhideColumnsAround() {
  const { sheet, view } = ctx();
  const { c1, c2 } = selRect(useUI.getState().sel);
  const from = c1 > 0 ? view.cols[c1 - 1] : 0;
  const to = c2 + 1 < view.cols.length ? view.cols[c2 + 1] : sheet.columns.length - 1;
  const ids = sheet.columns.slice(from, to + 1).filter((c) => c.hidden).map((c) => c.id);
  if (ids.length) store.setColumnsHidden(sheet, ids, false);
}

export function sortByColumn(vc: number, dir: 'asc' | 'desc') {
  const { sheet, view } = ctx();
  store.sortRows(sheet, view.cols[vc], dir);
}

// ─── фото ────────────────────────────────────────────────────────────────────

/** Вставить фото в ячейки начиная с (vr, vc) вниз по столбцу. */
export async function insertImages(files: File[] | Blob[], vr: number, vc: number) {
  const imgs = files.filter((f) => isImageFile(f as File));
  if (!imgs.length) {
    toast('Это не изображение. Подойдут JPG, PNG, WebP, GIF', { tone: 'error' });
    return;
  }
  const { sheet } = ctx();
  const ids: string[] = [];
  try {
    for (const f of imgs) ids.push(await importImage(f, (f as File).name));
  } catch (e) {
    toast(`Не удалось загрузить фото: ${e instanceof Error ? e.message : e}`, { tone: 'error' });
    return;
  }
  const view = store.view(sheet);
  const need = vr + ids.length - view.rows.length;
  store.transact(ids.length > 1 ? 'Вставка фото' : 'Вставка фото', () => {
    if (need > 0) store.insertRows(sheet, sheet.rowOrder.length, need);
    ids.forEach((id, i) => {
      const x = cellAt(vr + i, vc);
      if (x) store.patchCell(sheet, x.rowId, x.colId, (c) => ({ ...c, img: id }));
    });
    // невысокие строки увеличиваем, чтобы фото было видно
    const rowsIds = ids.map((_, i) => cellAt(vr + i, vc)?.rowId).filter(Boolean) as RowId[];
    const small = rowsIds.filter((id) => store.rowHeight(sheet, sheet.rows.get(id)) < 44);
    if (small.length) store.setRowsHeight(sheet, small, 72);
  });
  setSelection({ ar: vr, ac: vc, fr: vr + ids.length - 1, fc: vc });
  toast(ids.length > 1 ? `Добавлено фото: ${ids.length}` : 'Фото добавлено');
}

export function removeImage() {
  const { sheet } = ctx();
  store.transact('Удаление фото', () => {
    forEachSelected((x) => {
      if (x.row?.cells[x.colId]?.img) store.patchCell(sheet, x.rowId, x.colId, (c) => ({ ...c, img: undefined }));
    });
  });
}

// ─── история ─────────────────────────────────────────────────────────────────

export function undo() {
  if (useUI.getState().edit) cancelEdit();
  if (!store.canUndo()) return;
  store.undo();
  setSelection(useUI.getState().sel);
}

export function redo() {
  if (!store.canRedo()) return;
  store.redo();
  setSelection(useUI.getState().sel);
}

// ─── заполнение ──────────────────────────────────────────────────────────────

/** Копия ячейки со сдвигом формулы — для протягивания и вставки. */
export function shiftedCell(cell: Cell | undefined, dr: number, dc: number): Cell | undefined {
  if (!cell) return undefined;
  if (!cell.f) return { ...cell };
  return { ...cell, f: shiftFormula(cell.f, dr, dc) };
}

/** Ctrl+D: первая строка выделения — во все строки ниже. */
export function fillDown() {
  const { sheet } = ctx();
  const { r1, r2, c1, c2 } = selRect(useUI.getState().sel);
  if (r1 === r2) return;
  store.transact('Заполнить вниз', () => {
    for (let vc = c1; vc <= c2; vc++) {
      const src = cellAt(r1, vc);
      if (!src) continue;
      const cell = src.row?.cells[src.colId];
      for (let vr = r1 + 1; vr <= r2; vr++) {
        const dst = cellAt(vr, vc);
        if (!dst) continue;
        store.setCell(sheet, dst.rowId, dst.colId, shiftedCell(cell, dst.phys - src.phys, 0));
      }
    }
  });
}

export function fillRight() {
  const { sheet } = ctx();
  const { r1, r2, c1, c2 } = selRect(useUI.getState().sel);
  if (c1 === c2) return;
  store.transact('Заполнить вправо', () => {
    for (let vr = r1; vr <= r2; vr++) {
      const src = cellAt(vr, c1);
      if (!src) continue;
      const cell = src.row?.cells[src.colId];
      for (let vc = c1 + 1; vc <= c2; vc++) {
        const dst = cellAt(vr, vc);
        if (!dst) continue;
        store.setCell(sheet, dst.rowId, dst.colId, shiftedCell(cell, 0, dst.c - src.c));
      }
    }
  });
}

/**
 * Протягивание маркером: источник — текущее выделение, цель — до строки/столбца `to`.
 * Числовой ряд из 2+ значений продолжается арифметической прогрессией.
 */
export function fillTo(axis: 'row' | 'col', to: number) {
  const { sheet } = ctx();
  const s = useUI.getState().sel;
  const { r1, r2, c1, c2 } = selRect(s);
  store.transact('Протягивание', () => {
    if (axis === 'row') {
      const h = r2 - r1 + 1;
      for (let vc = c1; vc <= c2; vc++) {
        const series = numericSeries(Array.from({ length: h }, (_, i) => cellAt(r1 + i, vc)?.row?.cells[cellAt(r1 + i, vc)!.colId]));
        const down = to > r2;
        const range = down ? [r2 + 1, to] : [to, r1 - 1];
        for (let vr = range[0]; vr <= range[1]; vr++) {
          const dst = cellAt(vr, vc);
          if (!dst) continue;
          if (series) {
            const k = vr - r1;
            store.patchCell(sheet, dst.rowId, dst.colId, (c) => ({ ...c, v: series.start + series.step * k, f: undefined }));
            continue;
          }
          const srcVr = r1 + (((vr - r1) % h) + h) % h;
          const src = cellAt(srcVr, vc)!;
          store.setCell(sheet, dst.rowId, dst.colId, shiftedCell(src.row?.cells[src.colId], dst.phys - src.phys, 0));
        }
      }
      setSelection({ ...s, ar: Math.min(r1, to), fr: Math.max(r2, to) }, false);
    } else {
      const w = c2 - c1 + 1;
      for (let vr = r1; vr <= r2; vr++) {
        const right = to > c2;
        const range = right ? [c2 + 1, to] : [to, c1 - 1];
        for (let vc = range[0]; vc <= range[1]; vc++) {
          const dst = cellAt(vr, vc);
          if (!dst) continue;
          const srcVc = c1 + (((vc - c1) % w) + w) % w;
          const src = cellAt(vr, srcVc)!;
          store.setCell(sheet, dst.rowId, dst.colId, shiftedCell(src.row?.cells[src.colId], 0, dst.c - src.c));
        }
      }
      setSelection({ ...s, ac: Math.min(c1, to), fc: Math.max(c2, to) }, false);
    }
  });
}

function numericSeries(cells: (Cell | undefined)[]): { start: number; step: number } | null {
  if (cells.length < 2) return null;
  const nums = cells.map((c) => (c && !c.f && typeof c.v === 'number' ? c.v : null));
  if (nums.some((n) => n === null)) return null;
  const step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) if (Math.abs(nums[i]! - nums[i - 1]! - step) > 1e-9) return null;
  return { start: nums[0]!, step };
}

// ─── карточки ────────────────────────────────────────────────────────────────

export function openCard(vr: number) {
  const { view } = ctx();
  const rowId = view.rows[vr];
  if (rowId) useUI.getState().set({ dialog: { kind: 'card', rowId }, menu: null, edit: null });
}

/** Новая карточка = новая строка в конце листа со следующим свободным артикулом. */
export function newCard() {
  const { sheet } = ctx();
  const keyC = sheet.keyColId ? store.colIndexOf(sheet, sheet.keyColId) : -1;
  let nextSku: number | null = null;
  if (keyC >= 0) {
    let max = -Infinity;
    for (let i = 0; i < sheet.rowOrder.length; i++) {
      const v = store.cellValue(sheet.id, i, keyC);
      if (typeof v === 'number' && Number.isInteger(v)) max = Math.max(max, v);
    }
    if (Number.isFinite(max)) nextSku = max + 1;
  }
  // занимаем первую пустую строку в конце, если она есть
  let rowId: RowId | undefined;
  for (let i = sheet.rowOrder.length - 1; i >= 0; i--) {
    const row = sheet.rows.get(sheet.rowOrder[i]);
    if (!row) continue;
    const empty = Object.values(row.cells).every((c) => isEmptyCell({ ...c, st: undefined }));
    if (!empty) break;
    rowId = row.id;
  }
  store.transact('Новая карточка', () => {
    if (!rowId) rowId = store.insertRows(sheet, sheet.rowOrder.length, 1)[0];
    if (nextSku !== null && sheet.keyColId) store.patchCell(sheet, rowId!, sheet.keyColId, (c) => ({ ...c, v: nextSku! }));
  });
  const vr = store.view(sheet).rows.indexOf(rowId!);
  if (vr >= 0) setSelection({ ar: vr, ac: 0, fr: vr, fc: 0 });
  useUI.getState().set({ dialog: { kind: 'card', rowId: rowId! } });
}

// ─── прочее ──────────────────────────────────────────────────────────────────

export function rowLabel(sheet: Sheet, row: Row | undefined): string {
  if (!row || !sheet.keyColId) return '';
  const c = store.colIndexOf(sheet, sheet.keyColId);
  const i = sheet.rowOrder.indexOf(row.id);
  return c >= 0 && i >= 0 ? store.display(sheet, i, c).text : '';
}
