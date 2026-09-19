import { lettersToCol, normalizeLookalikes } from './a1';
import { ERROR_LITERALS, type ErrCode } from './values';

export interface RefEnd {
  /** Индекс столбца; отсутствует у ссылок на строки целиком (1:5). */
  c?: number;
  /** Индекс строки; отсутствует у ссылок на столбцы целиком (A:C). */
  r?: number;
  ca: boolean;
  ra: boolean;
}

interface Span {
  s: number;
  e: number;
}

export type Tok = Span &
  (
    | { k: 'num'; v: number }
    | { k: 'str'; v: string }
    | { k: 'bool'; v: boolean }
    | { k: 'err'; v: ErrCode }
    | { k: 'ref'; sheet?: string; a: RefEnd; b?: RefEnd }
    | { k: 'fn'; name: string }
    | { k: 'name'; name: string }
    | { k: 'op'; v: string }
    | { k: '(' }
    | { k: ')' }
    | { k: 'sep' }
  );

export class FormulaSyntaxError extends Error {}

const L = '[A-Za-zАВСЕНКМОРТХавсенкмортх]';
const CELL = `(\\$?)(${L}{1,3})(\\$?)(\\d+)`;
const RE_CELL_RANGE = new RegExp(`${CELL}(?::${CELL})?`, 'y');
const RE_COL_RANGE = new RegExp(`(\\$?)(${L}{1,3}):(\\$?)(${L}{1,3})`, 'y');
const RE_ROW_RANGE = /(\$?)(\d+):(\$?)(\d+)/y;
const RE_WORD = /[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_.]*/y;
const RE_NUM_DOT = /(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/y;
const RE_NUM_COMMA = /(\d+([.,]\d*)?|[.,]\d+)([eE][+-]?\d+)?/y;
const IDENT_CHAR = /[A-Za-zА-Яа-яЁё0-9_.(]/;

const MAX_COL = 16383;
const MAX_ROW = 1048575;

function cellEnd(dc: string, letters: string, dr: string, digits: string): RefEnd | null {
  const c = lettersToCol(normalizeLookalikes(letters));
  const r = parseInt(digits, 10) - 1;
  if (c > MAX_COL || r < 0 || r > MAX_ROW) return null;
  return { c, r, ca: dc === '$', ra: dr === '$' };
}

function tryRef(src: string, i: number): { a: RefEnd; b?: RefEnd; e: number } | null {
  RE_CELL_RANGE.lastIndex = i;
  let m = RE_CELL_RANGE.exec(src);
  if (m) {
    const end = RE_CELL_RANGE.lastIndex;
    if (end >= src.length || !IDENT_CHAR.test(src[end])) {
      const a = cellEnd(m[1], m[2], m[3], m[4]);
      const b = m[5] !== undefined ? cellEnd(m[5], m[6], m[7], m[8]) : undefined;
      if (a && b !== null) return { a, b, e: end };
    }
  }
  RE_COL_RANGE.lastIndex = i;
  m = RE_COL_RANGE.exec(src);
  if (m) {
    const end = RE_COL_RANGE.lastIndex;
    if (end >= src.length || !IDENT_CHAR.test(src[end])) {
      const c1 = lettersToCol(normalizeLookalikes(m[2]));
      const c2 = lettersToCol(normalizeLookalikes(m[4]));
      if (c1 <= MAX_COL && c2 <= MAX_COL) {
        return { a: { c: c1, ca: m[1] === '$', ra: false }, b: { c: c2, ca: m[3] === '$', ra: false }, e: end };
      }
    }
  }
  RE_ROW_RANGE.lastIndex = i;
  m = RE_ROW_RANGE.exec(src);
  if (m) {
    const end = RE_ROW_RANGE.lastIndex;
    const r1 = parseInt(m[2], 10) - 1;
    const r2 = parseInt(m[4], 10) - 1;
    if (r1 >= 0 && r2 >= 0 && (end >= src.length || !IDENT_CHAR.test(src[end]))) {
      return { a: { r: r1, ra: m[1] === '$', ca: false }, b: { r: r2, ra: m[3] === '$', ca: false }, e: end };
    }
  }
  return null;
}

/** Есть ли «;» вне строковых литералов — признак русской нотации (запятая = десятичный разделитель). */
export function usesSemicolons(src: string): boolean {
  let inStr = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') inStr = !inStr;
    else if (ch === ';' && !inStr) return true;
  }
  return false;
}

/**
 * Разбивает формулу на токены. `src` — полный текст, включая ведущий «=».
 * `commaDecimal` — запятая считается десятичным разделителем (русская нотация).
 */
export function tokenize(src: string, commaDecimal: boolean): Tok[] {
  const out: Tok[] = [];
  let i = src.startsWith('=') ? 1 : 0;
  const n = src.length;

  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\t' || ch === '\r' || ch === ' ') {
      i++;
      continue;
    }
    const s = i;

    // Строка
    if (ch === '"' || ch === '“' || ch === '«') {
      const close = ch === '"' ? '"' : ch === '“' ? '”' : '»';
      let v = '';
      i++;
      for (;;) {
        if (i >= n) throw new FormulaSyntaxError('Не закрыта кавычка');
        if (src[i] === close) {
          if (close === '"' && src[i + 1] === '"') {
            v += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        v += src[i++];
      }
      out.push({ k: 'str', v, s, e: i });
      continue;
    }

    // Литерал ошибки
    if (ch === '#') {
      const rest = src.slice(i).toUpperCase();
      const lit = ERROR_LITERALS.find(([text]) => rest.startsWith(text));
      if (!lit) throw new FormulaSyntaxError('Неизвестная ошибка');
      i += lit[0].length;
      out.push({ k: 'err', v: lit[1], s, e: i });
      continue;
    }

    // Ссылка на лист в кавычках: 'Лист 2'!A1
    if (ch === "'") {
      let name = '';
      i++;
      for (;;) {
        if (i >= n) throw new FormulaSyntaxError('Не закрыта кавычка в имени листа');
        if (src[i] === "'") {
          if (src[i + 1] === "'") {
            name += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        name += src[i++];
      }
      if (src[i] !== '!') throw new FormulaSyntaxError('После имени листа ожидается «!»');
      const ref = tryRef(src, i + 1);
      if (!ref) throw new FormulaSyntaxError('Ожидается ссылка на ячейку');
      i = ref.e;
      out.push({ k: 'ref', sheet: name, a: ref.a, b: ref.b, s, e: i });
      continue;
    }

    // Ссылки (включая строки целиком «1:3») пробуем раньше чисел и слов
    if (ch === '$' || /[0-9A-Za-zА-Яа-яЁё]/.test(ch)) {
      const ref = tryRef(src, i);
      if (ref) {
        i = ref.e;
        out.push({ k: 'ref', a: ref.a, b: ref.b, s, e: i });
        continue;
      }
    }

    // Число
    if (/[0-9]/.test(ch) || ((ch === '.' || (ch === ',' && commaDecimal)) && /[0-9]/.test(src[i + 1] ?? ''))) {
      const re = commaDecimal ? RE_NUM_COMMA : RE_NUM_DOT;
      re.lastIndex = i;
      const m = re.exec(src);
      if (!m) throw new FormulaSyntaxError('Некорректное число');
      i = re.lastIndex;
      out.push({ k: 'num', v: parseFloat(m[0].replace(',', '.')), s, e: i });
      continue;
    }

    // Слово: функция, имя листа, логическое значение или именованное значение
    if (/[A-Za-zА-Яа-яЁё_]/.test(ch)) {
      RE_WORD.lastIndex = i;
      const m = RE_WORD.exec(src)!;
      const word = m[0];
      i = RE_WORD.lastIndex;
      if (src[i] === '!') {
        const ref = tryRef(src, i + 1);
        if (!ref) throw new FormulaSyntaxError('Ожидается ссылка на ячейку');
        i = ref.e;
        out.push({ k: 'ref', sheet: word, a: ref.a, b: ref.b, s, e: i });
        continue;
      }
      let j = i;
      while (src[j] === ' ') j++;
      if (src[j] === '(') {
        out.push({ k: 'fn', name: word.toUpperCase(), s, e: i });
        continue;
      }
      const up = word.toUpperCase();
      if (up === 'TRUE' || up === 'ИСТИНА') {
        out.push({ k: 'bool', v: true, s, e: i });
        continue;
      }
      if (up === 'FALSE' || up === 'ЛОЖЬ') {
        out.push({ k: 'bool', v: false, s, e: i });
        continue;
      }
      out.push({ k: 'name', name: word, s, e: i });
      continue;
    }

    if (ch === '(') {
      out.push({ k: '(', s, e: ++i });
      continue;
    }
    if (ch === ')') {
      out.push({ k: ')', s, e: ++i });
      continue;
    }
    if (ch === ';' || (ch === ',' && !commaDecimal)) {
      out.push({ k: 'sep', s, e: ++i });
      continue;
    }
    if (ch === '<' || ch === '>') {
      const two = src.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') {
        i += 2;
        out.push({ k: 'op', v: two, s, e: i });
      } else {
        out.push({ k: 'op', v: ch, s, e: ++i });
      }
      continue;
    }
    if ('+-*/^&=%'.includes(ch)) {
      out.push({ k: 'op', v: ch, s, e: ++i });
      continue;
    }
    if (ch === '×') {
      out.push({ k: 'op', v: '*', s, e: ++i });
      continue;
    }
    throw new FormulaSyntaxError(`Неожиданный символ «${ch}»`);
  }
  return out;
}
