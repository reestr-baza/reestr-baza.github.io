import { colToLetters } from './a1';
import { tokenize, usesSemicolons, type RefEnd, type Tok } from './tokenize';

type RefTok = Extract<Tok, { k: 'ref' }>;
export interface Ref {
  sheet?: string;
  a: RefEnd;
  b?: RefEnd;
}

function tokensOf(src: string): Tok[] | null {
  const semi = usesSemicolons(src);
  try {
    return tokenize(src, semi);
  } catch {
    try {
      return tokenize(src, true);
    } catch {
      return null;
    }
  }
}

export function quoteSheetName(name: string): string {
  if (/^[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function endText(e: RefEnd): string {
  const col = e.c === undefined ? '' : `${e.ca ? '$' : ''}${colToLetters(e.c)}`;
  const row = e.r === undefined ? '' : `${e.ra ? '$' : ''}${e.r + 1}`;
  return col + row;
}

export function refText(ref: Ref): string {
  const prefix = ref.sheet ? `${quoteSheetName(ref.sheet)}!` : '';
  return prefix + endText(ref.a) + (ref.b ? `:${endText(ref.b)}` : '');
}

/**
 * Пересобирает формулу, заменяя каждую ссылку результатом `fn`.
 * `fn` возвращает новую ссылку, готовый текст (например, имя значения), `null`
 * (ссылка стала недействительной → #ССЫЛКА!) или `undefined` (оставить как есть).
 */
export function mapRefs(src: string, fn: (ref: Ref) => Ref | string | null | undefined): string {
  const toks = tokensOf(src);
  if (!toks) return src;
  let out = '';
  let pos = 0;
  let changed = false;
  for (const t of toks) {
    if (t.k !== 'ref') continue;
    const r = fn({ sheet: (t as RefTok).sheet, a: t.a, b: t.b });
    if (r === undefined) continue;
    out += src.slice(pos, t.s) + (r === null ? '#ССЫЛКА!' : typeof r === 'string' ? r : refText(r));
    pos = t.e;
    changed = true;
  }
  return changed ? out + src.slice(pos) : src;
}

/** Все ссылки формулы — для подсветки при редактировании. */
export function listRefs(src: string): (Ref & { s: number; e: number })[] {
  const toks = tokensOf(src);
  if (!toks) return [];
  return toks
    .filter((t): t is RefTok => t.k === 'ref')
    .map((t) => ({ sheet: t.sheet, a: t.a, b: t.b, s: t.s, e: t.e }));
}

function shiftEnd(e: RefEnd, dr: number, dc: number): RefEnd | null {
  const r = e.r === undefined || e.ra ? e.r : e.r + dr;
  const c = e.c === undefined || e.ca ? e.c : e.c + dc;
  if ((r !== undefined && r < 0) || (c !== undefined && c < 0)) return null;
  return { ...e, r, c };
}

/** Сдвиг относительных ссылок — при копировании, протягивании и сортировке. */
export function shiftFormula(src: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return src;
  return mapRefs(src, (ref) => {
    const a = shiftEnd(ref.a, dr, dc);
    const b = ref.b ? shiftEnd(ref.b, dr, dc) : undefined;
    if (!a || b === null) return null;
    return { ...ref, a, b };
  });
}

export type Axis = 'row' | 'col';

export interface StructuralChange {
  axis: Axis;
  /** Вставка: at, count > 0. Удаление: at, count < 0 (удаляется |count| позиций начиная с at). */
  at: number;
  count: number;
}

function adjustIndex(i: number, ch: StructuralChange): number | null {
  if (ch.count > 0) return i >= ch.at ? i + ch.count : i;
  const n = -ch.count;
  if (i < ch.at) return i;
  if (i >= ch.at + n) return i - n;
  return null;
}

/**
 * Подгоняет ссылки на лист после вставки/удаления строк или столбцов.
 * `touches(sheetName)` — относится ли ссылка к изменённому листу.
 * `anchoredRows` — формула столбца: её относительные строки привязаны к «своей» строке и не двигаются.
 */
export function adjustFormula(
  src: string,
  ch: StructuralChange,
  touches: (sheet: string | undefined) => boolean,
  anchoredRows = false,
): string {
  const key = ch.axis === 'row' ? 'r' : 'c';
  const absKey = ch.axis === 'row' ? 'ra' : 'ca';
  return mapRefs(src, (ref) => {
    if (!touches(ref.sheet)) return undefined;
    const fixed = (e: RefEnd) => anchoredRows && ch.axis === 'row' && !e[absKey];
    const i1 = ref.a[key];
    const i2 = ref.b ? ref.b[key] : undefined;
    if (i1 === undefined) return undefined; // столбец/строка целиком по другой оси
    if (!ref.b || i2 === undefined) {
      if (fixed(ref.a)) return undefined;
      const n = adjustIndex(i1, ch);
      if (n === null) return null;
      if (n === i1) return undefined;
      return { ...ref, a: { ...ref.a, [key]: n } };
    }
    // диапазон
    let lo = Math.min(i1, i2);
    let hi = Math.max(i1, i2);
    const loFixed = fixed(i1 <= i2 ? ref.a : ref.b);
    const hiFixed = fixed(i1 <= i2 ? ref.b : ref.a);
    if (ch.count > 0) {
      if (!loFixed && lo >= ch.at) lo += ch.count;
      if (!hiFixed && hi >= ch.at) hi += ch.count;
    } else {
      const n = -ch.count;
      const end = ch.at + n;
      if (!loFixed && !hiFixed && lo >= ch.at && hi < end) return null;
      if (!loFixed) lo = lo < ch.at ? lo : lo >= end ? lo - n : ch.at;
      if (!hiFixed) hi = hi < ch.at ? hi : hi >= end ? hi - n : ch.at - 1;
      if (hi < lo) return null;
    }
    const aIsLo = i1 <= i2;
    const a = { ...ref.a, [key]: aIsLo ? lo : hi };
    const b = { ...ref.b, [key]: aIsLo ? hi : lo };
    if (a[key] === i1 && b[key] === i2) return undefined;
    return { ...ref, a, b };
  });
}

/** Перестановка столбцов: старый индекс → новый. */
export function remapColumns(src: string, map: (c: number) => number, touches: (sheet: string | undefined) => boolean): string {
  return mapRefs(src, (ref) => {
    if (!touches(ref.sheet)) return undefined;
    const a = ref.a.c === undefined ? ref.a : { ...ref.a, c: map(ref.a.c) };
    const b = ref.b && ref.b.c !== undefined ? { ...ref.b, c: map(ref.b.c) } : ref.b;
    if (a.c === ref.a.c && b?.c === ref.b?.c) return undefined;
    // у диапазона после перестановки концы могут поменяться местами — нормализуем
    if (b && a.c !== undefined && b.c !== undefined && a.c > b.c) return { ...ref, a: { ...a, c: b.c }, b: { ...b, c: a.c } };
    return { ...ref, a, b };
  });
}

/** Переименование листа в ссылках других формул. */
export function renameSheetRefs(src: string, from: string, to: string): string {
  const low = from.toLocaleLowerCase('ru');
  return mapRefs(src, (ref) => (ref.sheet && ref.sheet.toLocaleLowerCase('ru') === low ? { ...ref, sheet: to } : undefined));
}
