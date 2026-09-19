import { numberToText } from '../formula/coerce';
import { parseInput } from '../model/format';
import type { Cell } from '../model/types';
import { selRect, useUI } from '../ui/state';
import { cellAt, ctx, insertImages, setSelection, shiftedCell, toast } from './actions';
import { requireEdit } from './editMode';
import { store } from './instance';

interface ClipCell {
  cell: Cell | undefined;
  phys: number;
  c: number;
  text: string;
}

interface InternalClip {
  id: string;
  sheetId: string;
  cut: boolean;
  cells: ClipCell[][];
  /** Откуда вырезали — чтобы очистить после вставки */
  sources: { rowId: string; colId: string }[];
}

let internal: InternalClip | null = null;

function escapeHtml(s: string) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function tsvField(s: string) {
  return /[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

/** Текст ячейки для внешних программ: числа без форматирования, с запятой. */
function plainText(vr: number, vc: number): string {
  const x = cellAt(vr, vc);
  if (!x) return '';
  const d = store.display(ctx().sheet, x.phys, x.c);
  if (typeof d.value === 'number' && d.style.nf?.k !== 'date' && d.style.nf?.k !== 'percent') return numberToText(d.value);
  return d.text;
}

export function handleCopy(e: ClipboardEvent, cut: boolean) {
  // вырезать в режиме просмотра нельзя: это изменение
  if (cut && !requireEdit()) {
    e.preventDefault();
    return;
  }
  const { sheet } = ctx();
  const sel = useUI.getState().sel;
  const { r1, r2, c1, c2 } = selRect(sel);
  const id = Math.random().toString(36).slice(2);
  const cells: ClipCell[][] = [];
  const sources: InternalClip['sources'] = [];
  const lines: string[] = [];
  let html = `<meta name="reestr-clip" content="${id}"><table>`;
  for (let vr = r1; vr <= r2; vr++) {
    const rowCells: ClipCell[] = [];
    const texts: string[] = [];
    html += '<tr>';
    for (let vc = c1; vc <= c2; vc++) {
      const x = cellAt(vr, vc);
      const text = plainText(vr, vc);
      texts.push(tsvField(text));
      const cell = x?.row?.cells[x.colId];
      // значение формулы столбца копируем как формулу, чтобы вставка вела себя как в Excel
      const f = x ? store.formulaAt(sheet, x.phys, x.c) : null;
      const effective: Cell | undefined = f && !cell?.f ? { ...cell, f } : cell;
      rowCells.push({ cell: effective, phys: x?.phys ?? 0, c: x?.c ?? 0, text });
      if (x) sources.push({ rowId: x.rowId, colId: x.colId });
      const st = x ? store.styleOf(sheet, x.col, cell) : {};
      const css = [
        st.b ? 'font-weight:bold' : '',
        st.i ? 'font-style:italic' : '',
        st.u ? 'text-decoration:underline' : '',
        st.bg ? `background:${st.bg}` : '',
        st.fg ? `color:${st.fg}` : '',
      ]
        .filter(Boolean)
        .join(';');
      html += `<td${css ? ` style="${css}"` : ''}>${escapeHtml(text)}</td>`;
    }
    html += '</tr>';
    cells.push(rowCells);
    lines.push(texts.join('\t'));
  }
  html += '</table>';
  e.clipboardData?.setData('text/plain', lines.join('\n'));
  e.clipboardData?.setData('text/html', html);
  e.preventDefault();
  internal = { id, sheetId: sheet.id, cut, cells, sources };
  useUI.getState().set({ copyRange: { r1, c1, r2, c2, cut } });
  const n = (r2 - r1 + 1) * (c2 - c1 + 1);
  toast(cut ? `Вырезано ячеек: ${n}` : n > 1 ? `Скопировано ячеек: ${n}` : 'Скопировано');
}

/** Разбор TSV с кавычками (формат Excel/Google Таблиц). */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  let quoted = false;
  const t = text.replace(/\r\n?/g, '\n');
  while (i < t.length) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
      i++;
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function parseHtmlTable(html: string): string[][] | null {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const out: string[][] = [];
  for (const tr of Array.from(table.querySelectorAll('tr'))) {
    const row: string[] = [];
    for (const td of Array.from(tr.querySelectorAll('td,th'))) {
      const text = (td as HTMLElement).innerText ?? td.textContent ?? '';
      row.push(text.replace(/ /g, ' ').trim());
      const span = Number(td.getAttribute('colspan') ?? 1);
      for (let k = 1; k < span; k++) row.push('');
    }
    out.push(row);
  }
  return out.length ? out : null;
}

/** Подготовить место: добавить строки/столбцы, если вставка выходит за край листа. */
function ensureSize(vr: number, vc: number, rows: number, cols: number): boolean {
  const { sheet, view } = ctx();
  const needRows = vr + rows - view.rows.length;
  const needCols = vc + cols - view.cols.length;
  if (needCols > 0) {
    if (sheet.columns.length + needCols > 200) {
      toast('Вставка не помещается: в листе максимум 200 столбцов', { tone: 'error' });
      return false;
    }
    store.insertColumns(sheet, sheet.columns.length, needCols);
  }
  if (needRows > 0) store.insertRows(sheet, sheet.rowOrder.length, needRows);
  return true;
}

export function handlePaste(e: ClipboardEvent) {
  const data = e.clipboardData;
  if (!data) return;
  if (!requireEdit()) {
    e.preventDefault();
    return;
  }
  const sel = useUI.getState().sel;
  const { r1, r2, c1, c2 } = selRect(sel);

  const files = Array.from(data.files ?? []).filter((f) => f.type.startsWith('image/'));
  if (files.length) {
    e.preventDefault();
    void insertImages(files, sel.ar, sel.ac);
    return;
  }

  const html = data.getData('text/html');
  const marker = /name="reestr-clip" content="([a-z0-9]+)"/.exec(html);
  if (marker && internal && internal.id === marker[1]) {
    e.preventDefault();
    pasteInternal(internal, r1, c1, r2, c2);
    return;
  }

  const text = data.getData('text/plain');
  let grid = html ? parseHtmlTable(html) : null;
  if (!grid) grid = text ? parseTsv(text) : null;
  if (!grid || !grid.length) return;
  e.preventDefault();
  pasteValues(grid, r1, c1, r2, c2);
}

function pasteValues(grid: string[][], r1: number, c1: number, r2: number, c2: number) {
  const { sheet } = ctx();
  const h = grid.length;
  const w = Math.max(...grid.map((r) => r.length));
  // одно значение в выделенный диапазон — заполнить весь диапазон
  const repeat = h === 1 && w === 1 && (r2 > r1 || c2 > c1);
  const rows = repeat ? r2 - r1 + 1 : h;
  const cols = repeat ? c2 - c1 + 1 : w;
  store.transact('Вставка', () => {
    if (!ensureSize(r1, c1, rows, cols)) return;
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        const raw = repeat ? grid[0][0] : (grid[i][j] ?? '');
        const x = cellAt(r1 + i, c1 + j);
        if (!x) continue;
        const parsed = parseInput(raw);
        store.patchCell(sheet, x.rowId, x.colId, (c) => {
          const next: Cell = { ...c };
          delete next.v;
          delete next.f;
          if (parsed.f && !(store.compile(parsed.f) instanceof Error)) next.f = parsed.f;
          else if (parsed.f) next.v = parsed.f;
          else if (parsed.v !== undefined) next.v = parsed.v;
          if (parsed.nf && !store.styleOf(sheet, x.col, c).nf) next.st = { ...next.st, nf: parsed.nf };
          return next;
        });
      }
  });
  setSelection({ ar: r1, ac: c1, fr: r1 + rows - 1, fc: c1 + cols - 1 }, false);
  if (rows * cols > 1) toast(`Вставлено ячеек: ${rows * cols}`);
}

function pasteInternal(clip: InternalClip, r1: number, c1: number, r2: number, c2: number) {
  const { sheet } = ctx();
  const h = clip.cells.length;
  const w = clip.cells[0]?.length ?? 0;
  const tileRows = Math.max(h, Math.floor((r2 - r1 + 1) / h) * h);
  const tileCols = Math.max(w, Math.floor((c2 - c1 + 1) / w) * w);
  store.transact(clip.cut ? 'Перемещение' : 'Вставка', () => {
    if (!ensureSize(r1, c1, tileRows, tileCols)) return;
    if (clip.cut && clip.sheetId === sheet.id) {
      for (const s of clip.sources) store.setCell(sheet, s.rowId, s.colId, undefined);
    } else if (clip.cut) {
      const src = store.sheets.get(clip.sheetId);
      if (src) for (const s of clip.sources) store.setCell(src, s.rowId, s.colId, undefined);
    }
    for (let i = 0; i < tileRows; i++)
      for (let j = 0; j < tileCols; j++) {
        const src = clip.cells[i % h][j % w];
        const x = cellAt(r1 + i, c1 + j);
        if (!x) continue;
        // при вырезании формулы не сдвигаются — ссылки остаются на те же ячейки
        const dr = clip.cut ? 0 : x.phys - src.phys;
        const dc = clip.cut ? 0 : x.c - src.c;
        store.setCell(sheet, x.rowId, x.colId, shiftedCell(src.cell, dr, dc));
      }
  });
  if (clip.cut) {
    internal = null;
    useUI.getState().set({ copyRange: null });
  }
  setSelection({ ar: r1, ac: c1, fr: r1 + tileRows - 1, fc: c1 + tileCols - 1 }, false);
}
