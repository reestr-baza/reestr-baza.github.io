import ExcelJS from 'exceljs';
import { colToLetters } from '../formula/a1';
import { toExcelFormula } from '../formula/excel';
import { isErr } from '../formula/values';
import { downloadBlob, safeFileName, stamp } from '../lib/download';
import { uid } from '../lib/ids';
import { parseInput } from '../model/format';
import { mergeBounds } from '../model/merges';
import type { Store } from '../model/store';
import { DEFAULT_COL_WIDTH, type Cell, type Column, type Currency, type NumFmt, type Row, type RowId } from '../model/types';
import { getStoredImage } from './images';

// ─── форматы ─────────────────────────────────────────────────────────────────

function decimalsOf(fmt: string): number {
  const m = /[.,](0+)/.exec(fmt.replace(/"[^"]*"/g, ''));
  return m ? m[1].length : 0;
}

export function numFmtFromExcel(fmt: string | undefined): NumFmt | undefined {
  if (!fmt || fmt === 'General') return undefined;
  const f = fmt.split(';')[0];
  const cur: [RegExp, Currency][] = [
    [/₽|р\.|руб|RUB/i, 'RUB'],
    [/¥|CN¥|CNY|元/i, 'CNY'],
    [/€|EUR/i, 'EUR'],
    [/\$|USD/i, 'USD'],
  ];
  for (const [re, c] of cur) if (re.test(f)) return { k: 'currency', c, d: decimalsOf(f) };
  if (f.includes('%')) return { k: 'percent', d: decimalsOf(f) };
  if (/(^|[^a-z])(d{1,4}|m{1,4}|y{2,4})([^a-z]|$)/i.test(f.replace(/"[^"]*"/g, '')) && !/[#0]/.test(f)) return { k: 'date' };
  if (f === '@') return { k: 'text' };
  if (/[#0]/.test(f)) return { k: 'number', d: decimalsOf(f) };
  return undefined;
}

export function numFmtToExcel(nf: NumFmt | undefined): string | undefined {
  if (!nf) return undefined;
  const dec = (d: number) => (d > 0 ? '.' + '0'.repeat(d) : '');
  switch (nf.k) {
    case 'number':
      return `#,##0${dec(nf.d)}`;
    case 'currency':
      if (nf.c === 'RUB') return `#,##0${dec(nf.d)} "₽"`;
      if (nf.c === 'CNY') return `"¥"#,##0${dec(nf.d)}`;
      if (nf.c === 'USD') return `"$"#,##0${dec(nf.d)}`;
      return `#,##0${dec(nf.d)} "€"`;
    case 'percent':
      return `0${dec(nf.d)}%`;
    case 'date':
      return 'dd.mm.yyyy';
    case 'text':
      return '@';
    default:
      return undefined;
  }
}

export function argbToHex(argb: string | undefined): string | undefined {
  if (!argb || !/^[0-9a-f]{8}$/i.test(argb)) return undefined;
  return '#' + argb.slice(2).toLowerCase();
}

function hexToArgb(hex: string): string {
  return 'FF' + hex.replace('#', '').toUpperCase().padStart(6, '0');
}

export { importXlsx, type ImportOptions, type ImportResult } from './xlsxImport';
import type { ImportOptions, ImportResult } from './xlsxImport';

// ─── CSV ─────────────────────────────────────────────────────────────────────

function detectDelimiter(line: string): string {
  const counts = [';', ',', '\t'].map((d) => [d, line.split(d).length] as const);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][0];
}

export function parseCsv(text: string): string[][] {
  const t = text.replace(/^﻿/, '').replace(/\r\n?/g, '\n');
  const delim = detectDelimiter(t.slice(0, t.indexOf('\n') > 0 ? t.indexOf('\n') : t.length));
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i++;
        } else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === '') quoted = true;
    else if (ch === delim) {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export async function importCsv(file: File, opts: ImportOptions): Promise<ImportResult> {
  const buf = await file.arrayBuffer();
  let text = new TextDecoder('utf-8').decode(buf);
  // файлы из русского Excel часто в windows-1251
  if (text.includes('�')) text = new TextDecoder('windows-1251').decode(buf);
  const grid = parseCsv(text);
  if (!grid.length) throw new Error('Файл пустой');
  const width = Math.min(200, Math.max(...grid.map((r) => r.length)));
  const hr = opts.headerRows === 'auto' ? 1 : opts.headerRows;
  const header = hr > 0 ? grid[hr - 1] : null;
  const columns: Column[] = Array.from({ length: width }, (_, c) => ({
    id: uid(6),
    name: header?.[c]?.trim() || colToLetters(c),
    w: DEFAULT_COL_WIDTH,
  }));
  const rows = new Map<RowId, Row>();
  const order: RowId[] = [];
  for (const line of grid.slice(hr)) {
    const id = uid(8);
    const cells: Record<string, Cell> = {};
    line.slice(0, width).forEach((raw, c) => {
      if (raw === '') return;
      const p = parseInput(raw);
      const cell: Cell = {};
      if (p.f) cell.v = raw;
      else if (p.v !== undefined) cell.v = p.v;
      if (p.nf) cell.st = { nf: p.nf };
      cells[columns[c].id] = cell;
    });
    rows.set(id, { id, cells });
    order.push(id);
  }
  const name = file.name.replace(/\.[^.]+$/, '').slice(0, 60) || 'Импорт';
  return {
    sheets: [{ id: uid(8), name, columns, rowOrder: order, rows, keyColId: columns[0]?.id, frozen: 1, density: 'S', filters: {} }],
    images: 0,
    rows: order.length,
    warnings: [],
    names: [],
  };
}

// ─── экспорт ─────────────────────────────────────────────────────────────────

/** Служебные листы выгрузки: при загрузке обратно они не становятся листами базы. */
export const PARAMS_SHEET = 'Параметры';
/** Цвет ссылок в выгрузке */
export const LINK_COLOR = '#1d5fc4';
export const PARAMS_HEAD = ['Имя', 'Значение', 'Пояснение'];
export const CARDS_SHEET = 'Карточки';
export const CARDS_HEAD = ['Лист', 'Строка', 'Артикул', 'Блок', 'Фото', 'Текст', 'Ширина', 'Высота', 'Оформление'];

/** Имя листа, каким его сохранит Excel: до 31 знака, без \ / ? * [ ] : */
export const excelSheetName = (name: string) => name.slice(0, 31).replace(/[\\/?*[\]:]/g, ' ');

/**
 * Фото в ячейку Excel: вписать с отступом 3px, привязка «перемещать и изменять вместе с ячейками».
 * Картинку добавляем в книгу для каждой ячейки заново: ExcelJS путает ссылки, если одну картинку
 * поставить в несколько мест (во второй ячейке оказывается чужое фото).
 */
function placeImage(ws: ExcelJS.Worksheet, imageId: number, img: { w: number; h: number }, row: number, col: number, colW: number, rowH: number) {
  const k = Math.min((colW - 6) / img.w, (rowH - 6) / img.h);
  const w = img.w * k;
  const h = img.h * k;
  const offX = (colW - w) / 2;
  const offY = (rowH - h) / 2;
  const EMU = 9525; // в одном пикселе
  ws.addImage(imageId, {
    tl: { nativeCol: col, nativeColOff: Math.round(offX * EMU), nativeRow: row, nativeRowOff: Math.round(offY * EMU) },
    br: { nativeCol: col, nativeColOff: Math.round((offX + w) * EMU), nativeRow: row, nativeRowOff: Math.round((offY + h) * EMU) },
    editAs: 'twoCell',
  } as never);
}

/** WebP в Excel не поддерживается — перекодируем в JPEG. */
async function imageForExcel(id: string, cache: Map<string, { buffer: ArrayBuffer; w: number; h: number } | null>) {
  if (cache.has(id)) return cache.get(id)!;
  const rec = await getStoredImage(id);
  if (!rec) {
    cache.set(id, null);
    return null;
  }
  const bmp = await createImageBitmap(rec.full);
  const canvas = document.createElement('canvas');
  canvas.width = bmp.width;
  canvas.height = bmp.height;
  const ctx = canvas.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bmp, 0, 0);
  bmp.close();
  const blob = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/jpeg', 0.88));
  const out = blob ? { buffer: await blob.arrayBuffer(), w: canvas.width, h: canvas.height } : null;
  cache.set(id, out);
  return out;
}

export async function exportXlsx(store: Store, scope: 'sheet' | 'all') {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Реестр';
  wb.created = new Date();
  // название базы — чтобы при загрузке файла обратно оно вернулось, а не стало именем файла
  wb.title = store.meta.title;
  const sheets = scope === 'all' ? store.sheetList() : [store.activeSheet];
  const imgCache = new Map<string, { buffer: ArrayBuffer; w: number; h: number } | null>();

  for (const sheet of sheets) {
    const ws = wb.addWorksheet(excelSheetName(sheet.name), {
      views: [{ state: 'frozen', xSplit: sheet.frozen, ySplit: 1 }],
    });
    ws.columns = sheet.columns.map((c) => ({ width: Math.max(4, (c.w - 5) / 7), hidden: !!c.hidden }));
    const head = ws.getRow(1);
    sheet.columns.forEach((c, i) => {
      const cell = head.getCell(i + 1);
      cell.value = c.name;
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4F5F2' } };
      cell.alignment = { vertical: 'middle' };
    });
    head.height = 22;
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: sheet.columns.length } };

    for (let i = 0; i < sheet.rowOrder.length; i++) {
      const row = sheet.rows.get(sheet.rowOrder[i]);
      if (!row) continue;
      const xr = ws.getRow(i + 2);
      const hPx = store.rowHeight(sheet, row);
      xr.height = hPx * 0.75;
      if (row.hidden) xr.hidden = true;
      for (let c = 0; c < sheet.columns.length; c++) {
        const col = sheet.columns[c];
        const cell = row.cells[col.id];
        const st = store.styleOf(sheet, col, cell);
        const xc = xr.getCell(c + 1);
        const formula = store.formulaAt(sheet, i, c);
        const value = store.cellValue(sheet.id, i, c);
        const result = isErr(value) ? undefined : value ?? undefined;
        if (formula) {
          xc.value = { formula: toExcelFormula(formula, 1), result: result as never } as ExcelJS.CellFormulaValue;
        } else if (cell?.href && cell.v !== undefined) {
          xc.value = { text: String(cell.v), hyperlink: cell.href };
          xc.font = { color: { argb: hexToArgb(LINK_COLOR) } };
        } else if (cell?.v !== undefined) {
          xc.value = st.nf?.k === 'date' && typeof cell.v === 'number' ? new Date(Date.UTC(1899, 11, 30) + cell.v * 86400000) : cell.v;
        }
        if (st.b || st.i || st.u || st.s || st.fg) {
          xc.font = { ...xc.font, bold: st.b, italic: st.i, underline: st.u, strike: st.s, color: st.fg ? { argb: hexToArgb(st.fg) } : xc.font?.color };
        }
        if (st.bg) xc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(st.bg) } };
        if (st.ha || st.va || st.wrap || cell?.img)
          xc.alignment = { horizontal: st.ha, vertical: st.va === 'middle' || !st.va ? 'middle' : st.va, wrapText: st.wrap };
        const fmt = numFmtToExcel(st.nf);
        if (fmt) xc.numFmt = fmt;
        if (cell?.note) xc.note = cell.note;

        if (cell?.img) {
          const img = await imageForExcel(cell.img, imgCache);
          if (!img) continue;
          placeImage(ws, wb.addImage({ buffer: img.buffer as never, extension: 'jpeg' }), img, i + 1, c, col.w, hPx);
        }
      }
    }
    // объединённые ячейки (строка 1 — шапка)
    for (const m of sheet.merges ?? []) {
      const b = mergeBounds(store, sheet, m);
      if (!b) continue;
      try {
        ws.mergeCells(b.p0 + 2, b.q0 + 1, b.p1 + 2, b.q1 + 1);
      } catch {
        /* пересекающееся объединение Excel не примет — пропускаем */
      }
    }
  }

  const cardRows = sheets.flatMap((sheet) =>
    sheet.rowOrder.map((id, i) => ({ sheet, row: sheet.rows.get(id), i })).filter((x) => x.row?.card?.length),
  );
  if (cardRows.length) {
    const cs = wb.addWorksheet(CARDS_SHEET, { views: [{ state: 'frozen', ySplit: 1 }] });
    cs.columns = [{ width: 18 }, { width: 9 }, { width: 14 }, { width: 7 }, { width: 22 }, { width: 50 }, { width: 9 }, { width: 9 }, { width: 30 }];
    const head = cs.getRow(1);
    head.values = CARDS_HEAD;
    head.font = { bold: true };
    let r = 2;
    for (const { sheet, row, i } of cardRows) {
      const keyC = sheet.keyColId ? store.colIndexOf(sheet, sheet.keyColId) : 0;
      const sku = store.display(sheet, i, Math.max(0, keyC)).text;
      for (const [b, block] of row!.card!.entries()) {
        const xr = cs.getRow(r);
        xr.values = [excelSheetName(sheet.name), i + 1, sku, b + 1, null, block.text ?? null, block.cs ?? 1, block.rs ?? 1, block.st ? JSON.stringify(block.st) : null];
        xr.getCell(6).alignment = { wrapText: true, vertical: 'top' };
        if (block.img) {
          const img = await imageForExcel(block.img, imgCache);
          if (img) {
            xr.height = 90;
            placeImage(cs, wb.addImage({ buffer: img.buffer as never, extension: 'jpeg' }), img, r - 1, 4, 22 * 7 + 5, 120);
          }
        }
        r++;
      }
    }
  }

  if (store.meta.names.length) {
    const ps = wb.addWorksheet(PARAMS_SHEET);
    ps.columns = [{ width: 28 }, { width: 16 }, { width: 40 }];
    ps.getRow(1).values = PARAMS_HEAD;
    ps.getRow(1).font = { bold: true };
    store.meta.names.forEach((n, i) => {
      ps.getRow(i + 2).values = [n.name, n.value, n.note ?? ''];
      wb.definedNames.add(`'${PARAMS_SHEET}'!$B$${i + 2}`, n.name);
    });
  }

  const buf = await wb.xlsx.writeBuffer();
  const name = scope === 'all' ? store.meta.title : `${store.meta.title} — ${store.activeSheet.name}`;
  downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeFileName(name)}_${stamp()}.xlsx`);
}

export function exportCsv(store: Store) {
  const sheet = store.activeSheet;
  const view = store.view(sheet);
  const esc = (s: string) => (/[;"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines = [view.cols.map((c) => esc(sheet.columns[c].name)).join(';')];
  for (const id of view.rows) {
    const r = view.rowIndex.get(id)!;
    lines.push(view.cols.map((c) => esc(store.display(sheet, r, c).text.replace(/ /g, ' '))).join(';'));
  }
  // BOM — чтобы Excel открыл кириллицу правильно
  downloadBlob(new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), `${safeFileName(sheet.name)}_${stamp()}.csv`);
}
