import ExcelJS from 'exceljs';
import { colToLetters } from '../formula/a1';
import { parseLooseNumber } from '../formula/coerce';
import { toRuFormula } from '../formula/excel';
import { dateToSerial } from '../formula/functions';
import { adjustFormula, listRefs, mapRefs, shiftFormula } from '../formula/refs';
import { uid } from '../lib/ids';
import { parseInput } from '../model/format';
import { DEFAULT_COL_WIDTH, type Cell, type CellStyle, type Column, type NamedValue, type Row, type RowId, type Sheet } from '../model/types';
import { imageSize, importImage } from './images';
import { argbToHex, numFmtFromExcel } from './xlsx';

export interface ImportOptions {
  /** Сколько строк сверху — заголовок. 'auto' — по закреплённым строкам файла. */
  headerRows: number | 'auto';
  onProgress?: (text: string) => void;
}

export interface ImportResult {
  sheets: Sheet[];
  images: number;
  rows: number;
  warnings: string[];
  /** Параметры из шапки (курс, доставка…), на которые ссылались формулы */
  names: NamedValue[];
}

type Anchor = { nativeRow?: number; row: number; nativeCol?: number; col: number };

function styleFromExcel(cell: ExcelJS.Cell): CellStyle | undefined {
  const st: CellStyle = {};
  const font = cell.font;
  if (font?.bold) st.b = true;
  if (font?.italic) st.i = true;
  if (font?.underline) st.u = true;
  if (font?.strike) st.s = true;
  const fg = argbToHex(font?.color?.argb);
  if (fg && fg !== '#000000') st.fg = fg;
  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const bg = argbToHex(fill.fgColor?.argb);
    if (bg && bg !== '#ffffff') st.bg = bg;
  }
  const al = cell.alignment;
  if (al?.horizontal === 'left' || al?.horizontal === 'center' || al?.horizontal === 'right') st.ha = al.horizontal;
  if (al?.vertical === 'top') st.va = 'top';
  if (al?.vertical === 'bottom') st.va = 'bottom';
  if (al?.wrapText) st.wrap = true;
  const nf = numFmtFromExcel(cell.numFmt);
  if (nf) st.nf = nf;
  return Object.keys(st).length ? st : undefined;
}

function richToText(v: unknown): string {
  if (v && typeof v === 'object') {
    const o = v as { richText?: { text: string }[]; texts?: { text: string }[] };
    if (Array.isArray(o.richText)) return o.richText.map((x) => x.text).join('');
    if (Array.isArray(o.texts)) return o.texts.map((x) => x.text).join('');
  }
  return v === null || v === undefined ? '' : String(v);
}

function cellFromExcel(cell: ExcelJS.Cell, convert: (excelFormula: string) => string): Cell | undefined {
  const v = cell.value;
  const out: Cell = {};
  const st = styleFromExcel(cell);
  if (st) out.st = st;
  const note = cell.note ? (typeof cell.note === 'string' ? cell.note : richToText(cell.note)).trim() : '';
  if (note) out.note = note;
  if (v === null || v === undefined) return st || note ? out : undefined;
  if (typeof v === 'number' || typeof v === 'boolean') out.v = v;
  else if (typeof v === 'string') {
    const p = parseInput(v);
    // текст из Excel остаётся текстом, кроме явных чисел (и не трогаем «00012»)
    out.v = typeof p.v === 'number' && !/^0\d/.test(v.trim()) ? p.v : v;
  } else if (v instanceof Date) {
    out.v = dateToSerial(new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate())) + (v.getUTCHours() * 3600 + v.getUTCMinutes() * 60) / 86400;
    if (!out.st?.nf) out.st = { ...out.st, nf: { k: 'date' } };
  } else if (typeof v === 'object') {
    const o = v as unknown as Record<string, unknown>;
    if ('formula' in o || 'sharedFormula' in o) {
      const f = cell.formula;
      if (f) out.f = convert(f);
      else if (o.result !== undefined && typeof o.result !== 'object') out.v = o.result as string | number | boolean;
    } else if ('richText' in o) {
      out.v = richToText(o);
    } else if ('hyperlink' in o) {
      const text = o.text as unknown;
      out.v = typeof text === 'string' ? text : richToText(text) || String(o.hyperlink);
      out.href = String(o.hyperlink);
    } else if ('error' in o) {
      out.v = String(o.error);
    } else if ('text' in o) {
      out.v = String(o.text);
    }
  }
  return out;
}

/** Текст подписи в шапке (не число, не пусто). */
function labelOf(v: unknown): string | null {
  if (v === null || v === undefined || typeof v === 'number' || typeof v === 'boolean' || v instanceof Date) return null;
  if (typeof v === 'object' && ('formula' in (v as object) || 'sharedFormula' in (v as object))) return null;
  const t = richToText(v).replace(/\s+/g, ' ').trim();
  if (!t || parseLooseNumber(t) !== null) return null;
  return t;
}

function scalarOf(v: unknown): number | string | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') return parseLooseNumber(v) ?? v;
  if (v && typeof v === 'object' && 'result' in (v as object)) {
    const r = (v as { result: unknown }).result;
    if (typeof r === 'number' || typeof r === 'string') return r;
  }
  return null;
}

/** Имя значения из подписи: «Курс» → «КУРС». Уникальное, не похожее на адрес ячейки. */
export function toValueName(raw: string, used: Set<string>): string {
  let n = raw
    .toLocaleUpperCase('ru')
    .replace(/[^A-ZА-ЯЁ0-9_]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
  if (!n) n = 'ЗНАЧЕНИЕ';
  if (/^\d/.test(n)) n = 'З_' + n;
  if (/^[A-Z]{1,3}\d+$/.test(n) || ['ИСТИНА', 'ЛОЖЬ', 'TRUE', 'FALSE'].includes(n)) n += '_';
  let out = n;
  let k = 2;
  while (used.has(out)) out = `${n}_${k++}`;
  used.add(out);
  return out;
}

/** Строки шапки: сколько строк закреплено в файле (1–4), иначе одна. */
export function detectHeaderRows(ws: ExcelJS.Worksheet): number {
  const v = ws.views?.[0] as { state?: string; ySplit?: number } | undefined;
  if (v?.state === 'frozen' && v.ySplit && v.ySplit >= 1 && v.ySplit <= 4) return v.ySplit;
  return 1;
}

async function pool<T>(items: T[], limit: number, fn: (x: T, i: number) => Promise<void>) {
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      await fn(items[i], i);
    }
  });
  await Promise.all(workers);
}

export async function importXlsx(file: File, opts: ImportOptions): Promise<ImportResult> {
  const wb = new ExcelJS.Workbook();
  opts.onProgress?.('Читаем файл…');
  await wb.xlsx.load(await file.arrayBuffer());
  const sheets: Sheet[] = [];
  const warnings: string[] = [];
  const names: NamedValue[] = [];
  const usedNames = new Set<string>();
  let imageCount = 0;
  let rowsTotal = 0;

  for (const ws of wb.worksheets) {
    if (ws.state === 'veryHidden') continue;
    const images = ws.getImages();
    let hasValues = false;
    ws.eachRow((row) => {
      if (hasValues) return;
      row.eachCell((c) => {
        if (c.value !== null && c.value !== undefined && c.value !== '') hasValues = true;
      });
    });
    if (!hasValues && !images.length) continue;

    const hr = opts.headerRows === 'auto' ? detectHeaderRows(ws) : Math.max(0, opts.headerRows);
    const rowCount = ws.rowCount;
    let colCount = Math.min(200, ws.columnCount || 0);
    if (ws.columnCount > 200) warnings.push(`Лист «${ws.name}»: взяты первые 200 столбцов из ${ws.columnCount}`);
    for (const img of images) {
      const tl = (img.range as unknown as { tl: Anchor }).tl;
      colCount = Math.max(colCount, Math.min(200, Math.floor(tl.nativeCol ?? tl.col) + 1));
    }
    opts.onProgress?.(`Лист «${ws.name}»: строк ${Math.max(0, rowCount - hr).toLocaleString('ru')}`);

    // названия столбцов — нижняя непустая подпись шапки
    const columns: Column[] = [];
    for (let c = 1; c <= Math.max(1, colCount); c++) {
      const wcol = ws.getColumn(c);
      let name = '';
      for (let r = hr; r >= 1 && !name; r--) name = labelOf(ws.getRow(r).getCell(c).value) ?? '';
      columns.push({
        id: uid(6),
        name: name || colToLetters(c - 1),
        w: wcol.width ? Math.round(Math.max(40, Math.min(900, wcol.width * 7 + 5))) : DEFAULT_COL_WIDTH,
        hidden: wcol.hidden || undefined,
      });
    }

    // значения в шапке, на которые ссылаются формулы ($J$1 — курс, $K$1 — доставка), становятся именованными
    const paramNames = new Map<string, string>();
    if (hr > 0) {
      const wanted = new Set<string>();
      for (let r = hr + 1; r <= rowCount; r++) {
        ws.getRow(r).eachCell((cell) => {
          const f = cell.formula;
          if (!f) return;
          for (const ref of listRefs('=' + f)) {
            if (ref.b || (ref.sheet && ref.sheet !== ws.name)) continue;
            if (ref.a.r !== undefined && ref.a.c !== undefined && ref.a.r < hr) wanted.add(ref.a.r + ':' + ref.a.c);
          }
        });
      }
      const byPosition = [...wanted].sort((p, q) => {
        const [pr, pc] = p.split(':').map(Number);
        const [qr, qc] = q.split(':').map(Number);
        return pr - qr || pc - qc;
      });
      for (const key of byPosition) {
        const [r, c] = key.split(':').map(Number);
        const row = ws.getRow(r + 1);
        const value = scalarOf(row.getCell(c + 1).value);
        if (value === null) continue;
        // подпись — ближайший текст слева: «Курс / Доставка» делим на части по порядку значений
        let k = c - 1;
        let index = 0;
        let label: string | null = null;
        while (k >= 0) {
          const v = row.getCell(k + 1).value;
          label = labelOf(v);
          if (label) break;
          if (scalarOf(v) !== null) index++;
          k--;
        }
        label ??= columns[c]?.name ?? colToLetters(c);
        const parts = label.split(/\s*[/,;]\s*|\s+и\s+/i).filter(Boolean);
        const raw = parts.length > 1 ? (parts[index] ?? `${label} ${index + 1}`) : index ? `${label} ${index + 1}` : label;
        const name = toValueName(raw, usedNames);
        paramNames.set(key, name);
        names.push({ name, value, note: `Из шапки листа «${ws.name}», ячейка ${colToLetters(c)}${r + 1}` });
      }
    }
    const convert = (f: string) => {
      let s = '=' + f.replace(/_xlfn\./gi, '').replace(/_xlws\./gi, '');
      if (paramNames.size) {
        s = mapRefs(s, (ref) => {
          if (ref.b || (ref.sheet && ref.sheet !== ws.name) || ref.a.r === undefined || ref.a.c === undefined) return undefined;
          return paramNames.get(ref.a.r + ':' + ref.a.c);
        });
      }
      if (hr) s = adjustFormula(s, { axis: 'row', at: 0, count: -hr }, () => true);
      return toRuFormula(s);
    };

    const rows = new Map<RowId, Row>();
    const order: RowId[] = [];
    const imageRows = new Set(images.map((i) => Math.floor((i.range as unknown as { tl: Anchor }).tl.nativeRow ?? 0)));
    const lastRow = Math.max(rowCount, ...[...imageRows].map((r) => r + 1));
    for (let r = hr + 1; r <= lastRow; r++) {
      const xr = ws.getRow(r);
      const id = uid(8);
      const cells: Record<string, Cell> = {};
      for (let c = 1; c <= columns.length; c++) {
        const cell = cellFromExcel(xr.getCell(c), convert);
        if (cell) cells[columns[c - 1].id] = cell;
      }
      const row: Row = { id, cells };
      if (xr.height) row.h = Math.round(Math.max(20, Math.min(600, (xr.height * 4) / 3)));
      if (xr.hidden) row.hidden = true;
      rows.set(id, row);
      order.push(id);
    }
    // хвост из пустых строк (только оформление) не нужен
    while (order.length > 1) {
      const last = rows.get(order[order.length - 1])!;
      const idx = order.length - 1 + hr;
      if (imageRows.has(idx) || Object.values(last.cells).some((c) => c.v !== undefined || c.f || c.note)) break;
      rows.delete(order.pop()!);
    }
    if (!order.length) {
      const id = uid(8);
      rows.set(id, { id, cells: {} });
      order.push(id);
    }

    // одинаковая формула во всех строках столбца → формула столбца (новые строки посчитаются сами)
    for (const col of columns) {
      let anchored: string | null = null;
      let withF = 0;
      let withV = 0;
      let same = true;
      order.forEach((id, i) => {
        const cell = rows.get(id)!.cells[col.id];
        if (cell?.f) {
          const a = shiftFormula(cell.f, -i, 0);
          if (anchored === null) anchored = a;
          else if (a !== anchored) same = false;
          withF++;
        } else if (cell?.v !== undefined) withV++;
      });
      const a = anchored as string | null;
      if (!same || !a || withF < 3 || withF < (withF + withV) * 0.9 || a.includes('#ССЫЛКА!')) continue;
      col.formula = a;
      for (const id of order) {
        const cell = rows.get(id)!.cells[col.id];
        if (cell?.f) delete cell.f;
      }
    }

    // фото: первое в ячейке — в ячейку, остальные из той же ячейки — в карточку товара
    const jobs = images
      .map((img) => {
        const tl = (img.range as unknown as { tl: Anchor }).tl;
        return { img, r0: Math.floor(tl.nativeRow ?? tl.row), c0: Math.floor(tl.nativeCol ?? tl.col) };
      })
      .sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
    const placed: { ri: number; c0: number; id: string; n: number }[] = [];
    let done = 0;
    await pool(jobs, 4, async ({ img, r0, c0 }, n) => {
      const ri = r0 - hr;
      const media = wb.getImage(Number(img.imageId)) as unknown as { buffer?: ArrayBuffer | Uint8Array; base64?: string; extension?: string };
      if (ri >= 0 && c0 < columns.length && media) {
        let bytes: Uint8Array | null = null;
        if (media.buffer) bytes = media.buffer instanceof Uint8Array ? media.buffer : new Uint8Array(media.buffer);
        else if (media.base64) bytes = Uint8Array.from(atob(media.base64.replace(/^data:[^,]+,/, '')), (ch) => ch.charCodeAt(0));
        if (bytes) {
          const ext = (media.extension ?? 'png').toLowerCase();
          try {
            const id = await importImage(new Blob([bytes as BlobPart], { type: `image/${ext === 'jpg' ? 'jpeg' : ext}` }), `${ws.name}-${colToLetters(c0)}${r0 + 1}.${ext}`);
            placed.push({ ri, c0, id, n });
          } catch {
            warnings.push(`Не удалось прочитать фото в ${colToLetters(c0)}${r0 + 1}`);
          }
        }
      }
      done++;
      if (done % 10 === 0 || done === jobs.length) opts.onProgress?.(`Лист «${ws.name}»: фото ${done} из ${jobs.length}`);
    });
    placed.sort((a, b) => a.n - b.n);
    // Excel часто хранит фото сжатыми превью — предупредим, почему в крупном виде они мягкие
    if (placed.length && !warnings.some((w) => w.startsWith('Фото в файле'))) {
      const sizes = (await Promise.all(placed.slice(0, 80).map((x) => imageSize(x.id)))).filter((x): x is { w: number; h: number } => !!x);
      if (sizes.length) {
        const med = (xs: number[]) => xs.sort((a, b) => a - b)[Math.floor(xs.length / 2)];
        const w = med(sizes.map((x) => x.w));
        const h = med(sizes.map((x) => x.h));
        if (Math.max(w, h) < 400) {
          warnings.unshift(`Фото в файле — маленькие превью, около ${w}×${h} px: такими их хранит Excel. В таблице они выглядят нормально, а для карточки загрузите оригиналы`);
        }
      }
    }
    const firstInCell = new Set<string>();
    for (const p of placed) {
      const row = rows.get(order[p.ri]);
      if (!row) continue;
      const colId = columns[p.c0].id;
      const key = p.ri + ':' + p.c0;
      if (!firstInCell.has(key)) {
        firstInCell.add(key);
        row.cells[colId] = { ...row.cells[colId], img: p.id };
      } else row.card = [...(row.card ?? []), { img: p.id }];
      if (!row.h || row.h < 60) row.h = 96;
      imageCount++;
    }

    const views = ws.views?.[0] as { state?: string; xSplit?: number } | undefined;
    const frozen = views?.state === 'frozen' && views.xSplit ? Math.min(views.xSplit, columns.length - 1) : 1;
    rowsTotal += order.length;
    sheets.push({
      id: uid(8),
      name: ws.name.slice(0, 60),
      columns,
      rowOrder: order,
      rows,
      keyColId: columns[0]?.id,
      frozen,
      density: images.length ? 'M' : 'S',
      filters: {},
    });
  }
  if (!sheets.length) throw new Error('В файле нет данных');
  return { sheets, images: imageCount, rows: rowsTotal, warnings, names };
}
