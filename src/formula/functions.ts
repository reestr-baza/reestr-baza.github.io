import { compareScalars, numberToText, parseLooseNumber, scalarOf, textEquals, toBool, toNumber, toText } from './coerce';
import type { EvalCtx } from './evaluate';
import { makeRange } from './range';
import type { Node } from './parse';
import { ERR, FErr, isErr, isRange, type RangeVal, type Scalar, type Value } from './values';

type Impl = (args: Value[], ctx: EvalCtx) => Value;
type LazyImpl = (nodes: Node[], ctx: EvalCtx, ev: (n: Node) => Value) => Value;

// ─── вспомогательные ─────────────────────────────────────────────────────────

/** Все значения аргументов подряд. В диапазонах текст и логические пропускаются (как в Excel). */
function* numbersOf(args: Value[]): Generator<number | FErr> {
  for (const a of args) {
    if (isRange(a)) {
      for (let r = 0; r < a.rows; r++)
        for (let c = 0; c < a.cols; c++) {
          const v = a.get(r, c);
          if (typeof v === 'number') yield v;
          else if (isErr(v)) yield v;
        }
    } else if (a !== null) {
      const n = toNumber(a);
      yield n;
    }
  }
}

function collectNumbers(args: Value[]): number[] | FErr {
  const out: number[] = [];
  for (const n of numbersOf(args)) {
    if (isErr(n)) return n;
    out.push(n);
  }
  return out;
}

function* scalarsOf(args: Value[]): Generator<Scalar> {
  for (const a of args) {
    if (isRange(a)) {
      for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) yield a.get(r, c);
    } else yield a;
  }
}

function num(v: Value | undefined, dflt?: number): number | FErr {
  if (v === undefined) return dflt ?? ERR.VALUE;
  const s = scalarOf(v);
  if (s === '' && dflt !== undefined) return dflt;
  return toNumber(s);
}

function txt(v: Value | undefined, dflt = ''): string | FErr {
  if (v === undefined) return dflt;
  return toText(scalarOf(v));
}

function asRange(v: Value, ctx: EvalCtx): RangeVal | FErr {
  if (isRange(v)) return v;
  if (isErr(v)) return v;
  return { kind: 'range', rows: 1, cols: 1, top: 0, left: 0, sheetId: ctx.sheetId, get: () => v };
}

function round(n: number, digits: number, mode: 'half' | 'up' | 'down'): number {
  const f = Math.pow(10, digits);
  const x = parseFloat((Math.abs(n) * f).toPrecision(15));
  const r = mode === 'half' ? Math.round(x) : mode === 'up' ? Math.ceil(x) : Math.floor(x);
  return (Math.sign(n) * r) / f;
}

// ─── условия СУММЕСЛИ/СЧЁТЕСЛИ ───────────────────────────────────────────────

type Matcher = (v: Scalar) => boolean;

function wildcardRegex(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'is');
}

export function makeMatcher(crit: Scalar): Matcher {
  if (typeof crit === 'number') return (v) => typeof v === 'number' && v === crit;
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => v === null || v === '';
  if (isErr(crit)) return (v) => isErr(v) && v.code === crit.code;
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(crit)!;
  const op = m[1] ?? '=';
  const rhs = m[2];
  const n = parseLooseNumber(rhs);
  if (n !== null && rhs.trim() !== '') {
    return (v) => {
      const x = typeof v === 'number' ? v : typeof v === 'string' ? parseLooseNumber(v) : null;
      if (x === null) return op === '<>';
      switch (op) {
        case '=':
          return x === n;
        case '<>':
          return x !== n;
        case '<':
          return x < n;
        case '>':
          return x > n;
        case '<=':
          return x <= n;
        default:
          return x >= n;
      }
    };
  }
  if (op === '=' || op === '<>') {
    if (rhs === '') {
      return op === '=' ? (v) => v === null || v === '' : (v) => !(v === null || v === '');
    }
    const hasWild = /[*?]/.test(rhs);
    const re = hasWild ? wildcardRegex(rhs) : null;
    const eq = (v: Scalar) => {
      const s = v === null ? '' : typeof v === 'string' ? v : String(toText(v));
      return re ? re.test(s) : textEquals(s, rhs);
    };
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const c = compareScalars(v, rhs);
    switch (op) {
      case '<':
        return c < 0;
      case '>':
        return c > 0;
      case '<=':
        return c <= 0;
      default:
        return c >= 0;
    }
  };
}

/** Общая логика ...ЕСЛИМН: пары (диапазон, условие); возвращает подходящие позиции. */
function matchPositions(pairs: Value[], ctx: EvalCtx): { r: number; c: number }[] | FErr {
  const ranges: RangeVal[] = [];
  const matchers: Matcher[] = [];
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    const rg = asRange(pairs[i], ctx);
    if (isErr(rg)) return rg;
    ranges.push(rg);
    matchers.push(makeMatcher(scalarOf(pairs[i + 1])));
  }
  if (!ranges.length) return ERR.VALUE;
  const { rows, cols } = ranges[0];
  for (const rg of ranges) if (rg.rows !== rows || rg.cols !== cols) return ERR.VALUE;
  const out: { r: number; c: number }[] = [];
  for (let r = 0; r < rows; r++)
    for (let c = 0; c < cols; c++) {
      let ok = true;
      for (let k = 0; k < ranges.length && ok; k++) ok = matchers[k](ranges[k].get(r, c));
      if (ok) out.push({ r, c });
    }
  return out;
}

function aggregateIfs(sumRange: Value, pairs: Value[], ctx: EvalCtx, fold: (xs: number[]) => Value): Value {
  const sr = asRange(sumRange, ctx);
  if (isErr(sr)) return sr;
  const pos = matchPositions(pairs, ctx);
  if (isErr(pos)) return pos;
  const xs: number[] = [];
  for (const { r, c } of pos) {
    const v = sr.get(r, c);
    if (isErr(v)) return v;
    if (typeof v === 'number') xs.push(v);
  }
  return fold(xs);
}

// ─── даты ────────────────────────────────────────────────────────────────────

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

export function dateToSerial(d: Date): number {
  return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EPOCH) / DAY_MS;
}

export function serialToDate(serial: number): Date {
  return new Date(EPOCH + Math.floor(serial) * DAY_MS);
}

function ymd(serial: number) {
  const d = serialToDate(serial);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

function makeSerial(y: number, m: number, d: number): number {
  return (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
}

// ─── формат ТЕКСТ() ──────────────────────────────────────────────────────────

export function formatByPattern(value: number, pattern: string): string {
  const p = pattern.trim();
  if (/[дмгdmy]/i.test(p) && !/[0#]/.test(p)) {
    const { y, m, d } = ymd(value);
    const pad = (n: number) => String(n).padStart(2, '0');
    return p
      .replace(/гггг|yyyy/gi, String(y))
      .replace(/гг|yy/gi, pad(y % 100))
      .replace(/мм|mm/gi, pad(m))
      .replace(/дд|dd/gi, pad(d));
  }
  const percent = p.endsWith('%');
  let v = percent ? value * 100 : value;
  const body = p.replace(/%$/, '');
  const decMatch = /[.,](0+)$/.exec(body);
  const decimals = decMatch ? decMatch[1].length : 0;
  const intPart = decMatch ? body.slice(0, decMatch.index) : body;
  const group = /[ , ]/.test(intPart.replace(/^[^0#]*/, '').replace(/[0#]$/, '')) || /# ?##0/.test(intPart);
  v = round(v, decimals, 'half');
  const [ip, fp] = Math.abs(v).toFixed(decimals).split('.');
  const intText = group ? ip.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') : ip;
  return `${v < 0 ? '-' : ''}${intText}${fp ? ',' + fp : ''}${percent ? '%' : ''}`;
}

// ─── поиск ───────────────────────────────────────────────────────────────────

function lookupKey(v: Scalar): string {
  if (typeof v === 'number') return 'n:' + v;
  if (typeof v === 'string') return 's:' + v.toLocaleLowerCase('ru');
  if (typeof v === 'boolean') return 'b:' + v;
  return 'e';
}

/** Точный поиск значения в одномерном диапазоне (столбец или строка). Возвращает индекс или -1. */
function exactFind(rg: RangeVal, key: Scalar, ctx: EvalCtx, vertical: boolean): number {
  const len = vertical ? rg.rows : rg.cols;
  if (typeof key === 'string' && /[*?]/.test(key)) {
    const re = wildcardRegex(key);
    for (let i = 0; i < len; i++) {
      const v = vertical ? rg.get(i, 0) : rg.get(0, i);
      if (typeof v === 'string' && re.test(v)) return i;
    }
    return -1;
  }
  if (vertical && len > 64 && ctx.host.lookupIndex) {
    const idx = ctx.host.lookupIndex(rg, 0);
    return idx.get(lookupKey(key)) ?? -1;
  }
  for (let i = 0; i < len; i++) {
    const v = vertical ? rg.get(i, 0) : rg.get(0, i);
    if (v !== null && compareScalars(v, key) === 0 && typeof v === typeof key) return i;
    if (typeof v === 'number' && typeof key === 'string' && parseLooseNumber(key) === v) return i;
    if (typeof v === 'string' && typeof key === 'number' && parseLooseNumber(v) === key) return i;
  }
  return -1;
}

/** Приближённый поиск (отсортированные данные): последний элемент ≤ ключа. */
function approxFind(rg: RangeVal, key: Scalar, vertical: boolean, descending = false): number {
  const len = vertical ? rg.rows : rg.cols;
  let found = -1;
  for (let i = 0; i < len; i++) {
    const v = vertical ? rg.get(i, 0) : rg.get(0, i);
    if (v === null) continue;
    const c = compareScalars(v, key);
    if (descending ? c >= 0 : c <= 0) found = i;
    else break;
  }
  return found;
}

export function buildLookupIndex(rg: RangeVal, col: number): Map<string, number> {
  const map = new Map<string, number>();
  for (let r = 0; r < rg.rows; r++) {
    const v = rg.get(r, col);
    if (v === null) continue;
    const k = lookupKey(v);
    if (!map.has(k)) map.set(k, r);
    if (typeof v === 'string') {
      const n = parseLooseNumber(v);
      if (n !== null && !map.has('n:' + n)) map.set('n:' + n, r);
    } else if (typeof v === 'number') {
      const s = 's:' + numberToText(v).toLocaleLowerCase('ru');
      if (!map.has(s)) map.set(s, r);
    }
  }
  return map;
}

// ─── реестр функций ──────────────────────────────────────────────────────────

const sum = (xs: number[]) => {
  let s = 0;
  for (const x of xs) s += x;
  return parseFloat(s.toPrecision(15));
};
const avg = (xs: number[]) => (xs.length ? sum(xs) / xs.length : ERR.DIV0);
const min = (xs: number[]) => (xs.length ? Math.min(...xs) : 0);
const max = (xs: number[]) => (xs.length ? Math.max(...xs) : 0);

function numFold(fold: (xs: number[]) => Value): Impl {
  return (args) => {
    const xs = collectNumbers(args);
    return isErr(xs) ? xs : fold(xs);
  };
}

function textFn(fn: (s: string, args: Value[]) => Value): Impl {
  return (args) => {
    const s = txt(args[0]);
    return isErr(s) ? s : fn(s, args);
  };
}

function subtotal(code: number, ranges: Value[], ctx: EvalCtx): Value {
  const ignoreManual = code > 100;
  const base = code % 100;
  const xs: number[] = [];
  let counta = 0;
  for (const a of ranges) {
    if (!isRange(a)) continue;
    for (let r = 0; r < a.rows; r++) {
      const h = ctx.host.rowHidden(a.sheetId, a.top + r);
      if (h.filtered || (ignoreManual && h.manual)) continue;
      for (let c = 0; c < a.cols; c++) {
        const v = a.get(r, c);
        if (isErr(v)) return v;
        if (v !== null && v !== '') counta++;
        if (typeof v === 'number') xs.push(v);
      }
    }
  }
  switch (base) {
    case 1:
      return avg(xs);
    case 2:
      return xs.length;
    case 3:
      return counta;
    case 4:
      return max(xs);
    case 5:
      return min(xs);
    case 6:
      return xs.reduce((p, x) => p * x, xs.length ? 1 : 0);
    case 9:
      return sum(xs);
    default:
      return ERR.VALUE;
  }
}

export const FUNCTIONS: Record<string, Impl> = {
  SUM: numFold(sum),
  AVERAGE: numFold(avg),
  MIN: numFold(min),
  MAX: numFold(max),
  PRODUCT: numFold((xs) => (xs.length ? xs.reduce((p, x) => p * x, 1) : 0)),
  MEDIAN: numFold((xs) => {
    if (!xs.length) return ERR.NUM;
    const s = [...xs].sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }),
  COUNT: (args) => {
    let n = 0;
    for (const a of args) {
      if (isRange(a)) {
        for (const v of scalarsOf([a])) if (typeof v === 'number') n++;
      } else if (typeof a === 'number' || (typeof a === 'string' && parseLooseNumber(a) !== null)) n++;
    }
    return n;
  },
  COUNTA: (args) => {
    let n = 0;
    for (const v of scalarsOf(args)) if (v !== null && v !== '') n++;
    return n;
  },
  COUNTBLANK: (args) => {
    let n = 0;
    for (const v of scalarsOf(args)) if (v === null || v === '') n++;
    return n;
  },
  ABS: (args) => {
    const n = num(args[0]);
    return isErr(n) ? n : Math.abs(n);
  },
  INT: (args) => {
    const n = num(args[0]);
    return isErr(n) ? n : Math.floor(n);
  },
  MOD: (args) => {
    const a = num(args[0]);
    const b = num(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (b === 0) return ERR.DIV0;
    return a - b * Math.floor(a / b);
  },
  POWER: (args) => {
    const a = num(args[0]);
    const b = num(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    const r = Math.pow(a, b);
    return Number.isFinite(r) ? r : ERR.NUM;
  },
  SQRT: (args) => {
    const n = num(args[0]);
    if (isErr(n)) return n;
    return n < 0 ? ERR.NUM : Math.sqrt(n);
  },
  ROUND: (args) => {
    const n = num(args[0]);
    const d = num(args[1], 0);
    if (isErr(n)) return n;
    if (isErr(d)) return d;
    return round(n, Math.trunc(d), 'half');
  },
  ROUNDUP: (args) => {
    const n = num(args[0]);
    const d = num(args[1], 0);
    if (isErr(n)) return n;
    if (isErr(d)) return d;
    return round(n, Math.trunc(d), 'up');
  },
  ROUNDDOWN: (args) => {
    const n = num(args[0]);
    const d = num(args[1], 0);
    if (isErr(n)) return n;
    if (isErr(d)) return d;
    return round(n, Math.trunc(d), 'down');
  },
  CEILING: (args) => {
    const n = num(args[0]);
    const s = num(args[1], 1);
    if (isErr(n)) return n;
    if (isErr(s)) return s;
    return s === 0 ? 0 : Math.ceil(n / s) * s;
  },
  FLOOR: (args) => {
    const n = num(args[0]);
    const s = num(args[1], 1);
    if (isErr(n)) return n;
    if (isErr(s)) return s;
    return s === 0 ? ERR.DIV0 : Math.floor(n / s) * s;
  },
  SUMPRODUCT: (args, ctx) => {
    const ranges: RangeVal[] = [];
    for (const a of args) {
      const r = asRange(a, ctx);
      if (isErr(r)) return r;
      ranges.push(r);
    }
    if (!ranges.length) return ERR.VALUE;
    const { rows, cols } = ranges[0];
    if (ranges.some((r) => r.rows !== rows || r.cols !== cols)) return ERR.VALUE;
    let s = 0;
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const rg of ranges) {
          const v = rg.get(r, c);
          if (isErr(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        s += p;
      }
    return s;
  },
  SUMIF: (args, ctx) => aggregateIfs(args[2] ?? args[0], [args[0], args[1]], ctx, sum),
  SUMIFS: (args, ctx) => aggregateIfs(args[0], args.slice(1), ctx, sum),
  AVERAGEIF: (args, ctx) => aggregateIfs(args[2] ?? args[0], [args[0], args[1]], ctx, avg),
  AVERAGEIFS: (args, ctx) => aggregateIfs(args[0], args.slice(1), ctx, avg),
  MAXIFS: (args, ctx) => aggregateIfs(args[0], args.slice(1), ctx, max),
  MINIFS: (args, ctx) => aggregateIfs(args[0], args.slice(1), ctx, min),
  COUNTIF: (args, ctx) => {
    const pos = matchPositions([args[0], args[1]], ctx);
    return isErr(pos) ? pos : pos.length;
  },
  COUNTIFS: (args, ctx) => {
    const pos = matchPositions(args, ctx);
    return isErr(pos) ? pos : pos.length;
  },
  SUBTOTAL: (args, ctx) => {
    const code = num(args[0]);
    if (isErr(code)) return code;
    return subtotal(Math.trunc(code), args.slice(1), ctx);
  },

  AND: (args) => {
    let any = false;
    for (const v of scalarsOf(args)) {
      if (v === null || typeof v === 'string') continue;
      const b = toBool(v);
      if (isErr(b)) return b;
      any = true;
      if (!b) return false;
    }
    return any ? true : ERR.VALUE;
  },
  OR: (args) => {
    let any = false;
    for (const v of scalarsOf(args)) {
      if (v === null || typeof v === 'string') continue;
      const b = toBool(v);
      if (isErr(b)) return b;
      any = true;
      if (b) return true;
    }
    return any ? false : ERR.VALUE;
  },
  NOT: (args) => {
    const b = toBool(scalarOf(args[0] ?? null));
    return isErr(b) ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,
  ISBLANK: (args) => {
    const v = scalarOf(args[0] ?? null);
    return v === null;
  },
  ISNUMBER: (args) => typeof scalarOf(args[0] ?? null) === 'number',
  ISTEXT: (args) => typeof scalarOf(args[0] ?? null) === 'string',
  ISERROR: (args) => isErr(scalarOf(args[0] ?? null)),
  NA: () => ERR.NA,

  CONCATENATE: (args) => {
    let s = '';
    for (const v of scalarsOf(args)) {
      const t = toText(v);
      if (isErr(t)) return t;
      s += t;
    }
    return s;
  },
  TEXTJOIN: (args) => {
    const sep = txt(args[0]);
    if (isErr(sep)) return sep;
    const skip = toBool(scalarOf(args[1] ?? true));
    if (isErr(skip)) return skip;
    const parts: string[] = [];
    for (const v of scalarsOf(args.slice(2))) {
      const t = toText(v);
      if (isErr(t)) return t;
      if (skip && t === '') continue;
      parts.push(t);
    }
    return parts.join(sep);
  },
  LEFT: textFn((s, a) => {
    const n = num(a[1], 1);
    return isErr(n) ? n : n < 0 ? ERR.VALUE : s.slice(0, n);
  }),
  RIGHT: textFn((s, a) => {
    const n = num(a[1], 1);
    return isErr(n) ? n : n < 0 ? ERR.VALUE : n === 0 ? '' : s.slice(-n);
  }),
  MID: textFn((s, a) => {
    const start = num(a[1]);
    const len = num(a[2]);
    if (isErr(start)) return start;
    if (isErr(len)) return len;
    if (start < 1 || len < 0) return ERR.VALUE;
    return s.substr(start - 1, len);
  }),
  LEN: textFn((s) => s.length),
  UPPER: textFn((s) => s.toLocaleUpperCase('ru')),
  LOWER: textFn((s) => s.toLocaleLowerCase('ru')),
  PROPER: textFn((s) => s.toLocaleLowerCase('ru').replace(/(^|[^\p{L}])(\p{L})/gu, (_m, p, ch) => p + ch.toLocaleUpperCase('ru'))),
  TRIM: textFn((s) => s.trim().replace(/ {2,}/g, ' ')),
  SUBSTITUTE: textFn((s, a) => {
    const from = txt(a[1]);
    const to = txt(a[2]);
    if (isErr(from)) return from;
    if (isErr(to)) return to;
    if (from === '') return s;
    if (a[3] === undefined) return s.split(from).join(to);
    const nth = num(a[3]);
    if (isErr(nth)) return nth;
    let idx = -1;
    for (let k = 0; k < nth; k++) {
      idx = s.indexOf(from, idx + 1);
      if (idx < 0) return s;
    }
    return s.slice(0, idx) + to + s.slice(idx + from.length);
  }),
  REPLACE: textFn((s, a) => {
    const start = num(a[1]);
    const len = num(a[2]);
    const rep = txt(a[3]);
    if (isErr(start)) return start;
    if (isErr(len)) return len;
    if (isErr(rep)) return rep;
    return s.slice(0, start - 1) + rep + s.slice(start - 1 + len);
  }),
  FIND: (args) => {
    const needle = txt(args[0]);
    const hay = txt(args[1]);
    const start = num(args[2], 1);
    if (isErr(needle)) return needle;
    if (isErr(hay)) return hay;
    if (isErr(start)) return start;
    const i = hay.indexOf(needle, start - 1);
    return i < 0 ? ERR.VALUE : i + 1;
  },
  SEARCH: (args) => {
    const needle = txt(args[0]);
    const hay = txt(args[1]);
    const start = num(args[2], 1);
    if (isErr(needle)) return needle;
    if (isErr(hay)) return hay;
    if (isErr(start)) return start;
    const re = wildcardRegex(needle);
    const src = re.source.slice(1, -1);
    const m = new RegExp(src, 'is').exec(hay.slice(start - 1));
    return m ? m.index + start : ERR.VALUE;
  },
  REPT: textFn((s, a) => {
    const n = num(a[1]);
    return isErr(n) ? n : n < 0 ? ERR.VALUE : s.repeat(Math.min(n, 10000));
  }),
  EXACT: (args) => {
    const a = txt(args[0]);
    const b = txt(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return a === b;
  },
  VALUE: (args) => {
    const v = scalarOf(args[0] ?? null);
    if (typeof v === 'number') return v;
    if (typeof v !== 'string') return ERR.VALUE;
    const n = parseLooseNumber(v.replace(/[₽¥$€]|руб\.?|р\./gi, ''));
    return n === null ? ERR.VALUE : n;
  },
  TEXT: (args) => {
    const v = scalarOf(args[0] ?? null);
    const pattern = txt(args[1]);
    if (isErr(pattern)) return pattern;
    if (isErr(v)) return v;
    if (typeof v === 'string') {
      const n = parseLooseNumber(v);
      return n === null ? v : formatByPattern(n, pattern);
    }
    const n = toNumber(v);
    return isErr(n) ? n : formatByPattern(n, pattern);
  },

  VLOOKUP: (args, ctx) => {
    const key = scalarOf(args[0] ?? null);
    const rg = args[1] === undefined ? ERR.VALUE : asRange(args[1], ctx);
    const colN = num(args[2]);
    if (isErr(key)) return key;
    if (isErr(rg)) return rg;
    if (isErr(colN)) return colN;
    const approx = args[3] === undefined ? true : toBool(scalarOf(args[3]));
    if (isErr(approx)) return approx;
    if (colN < 1 || colN > rg.cols) return ERR.REF;
    const first = makeRange(ctx, rg.sheetId, rg.top, rg.left, rg.top + rg.rows - 1, rg.left);
    const i = approx ? approxFind(first, key, true) : exactFind(first, key, ctx, true);
    return i < 0 ? ERR.NA : rg.get(i, colN - 1);
  },
  HLOOKUP: (args, ctx) => {
    const key = scalarOf(args[0] ?? null);
    const rg = args[1] === undefined ? ERR.VALUE : asRange(args[1], ctx);
    const rowN = num(args[2]);
    if (isErr(key)) return key;
    if (isErr(rg)) return rg;
    if (isErr(rowN)) return rowN;
    const approx = args[3] === undefined ? true : toBool(scalarOf(args[3]));
    if (isErr(approx)) return approx;
    if (rowN < 1 || rowN > rg.rows) return ERR.REF;
    const i = approx ? approxFind(rg, key, false) : exactFind(rg, key, ctx, false);
    return i < 0 ? ERR.NA : rg.get(rowN - 1, i);
  },
  MATCH: (args, ctx) => {
    const key = scalarOf(args[0] ?? null);
    const rg = args[1] === undefined ? ERR.VALUE : asRange(args[1], ctx);
    const type = num(args[2], 1);
    if (isErr(key)) return key;
    if (isErr(rg)) return rg;
    if (isErr(type)) return type;
    if (rg.rows !== 1 && rg.cols !== 1) return ERR.NA;
    const vertical = rg.cols === 1;
    const i = type === 0 ? exactFind(rg, key, ctx, vertical) : approxFind(rg, key, vertical, type < 0);
    return i < 0 ? ERR.NA : i + 1;
  },
  INDEX: (args, ctx) => {
    const rg = args[0] === undefined ? ERR.VALUE : asRange(args[0], ctx);
    if (isErr(rg)) return rg;
    let r = num(args[1], 0);
    let c = num(args[2], 0);
    if (isErr(r)) return r;
    if (isErr(c)) return c;
    if (rg.rows === 1 && args[2] === undefined) {
      c = r;
      r = 1;
    }
    if (r === 0 && rg.rows === 1) r = 1;
    if (c === 0 && rg.cols === 1) c = 1;
    if (r < 1 || c < 1 || r > rg.rows || c > rg.cols) return ERR.REF;
    return rg.get(r - 1, c - 1);
  },
  XLOOKUP: (args, ctx) => {
    const key = scalarOf(args[0] ?? null);
    const look = args[1] === undefined ? ERR.VALUE : asRange(args[1], ctx);
    const ret = args[2] === undefined ? ERR.VALUE : asRange(args[2], ctx);
    if (isErr(key)) return key;
    if (isErr(look)) return look;
    if (isErr(ret)) return ret;
    const vertical = look.cols === 1;
    const i = exactFind(look, key, ctx, vertical);
    if (i < 0) return args[3] !== undefined ? scalarOf(args[3]) : ERR.NA;
    return vertical ? ret.get(i, 0) : ret.get(0, i);
  },
  ROWS: (args, ctx) => {
    const rg = args[0] === undefined ? ERR.VALUE : asRange(args[0], ctx);
    return isErr(rg) ? rg : rg.rows;
  },
  COLUMNS: (args, ctx) => {
    const rg = args[0] === undefined ? ERR.VALUE : asRange(args[0], ctx);
    return isErr(rg) ? rg : rg.cols;
  },

  TODAY: () => dateToSerial(new Date()),
  NOW: () => {
    const d = new Date();
    return dateToSerial(d) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
  },
  DATE: (args) => {
    const y = num(args[0]);
    const m = num(args[1]);
    const d = num(args[2]);
    if (isErr(y)) return y;
    if (isErr(m)) return m;
    if (isErr(d)) return d;
    return makeSerial(y < 1900 ? y + 1900 : y, m, d);
  },
  YEAR: (args) => {
    const n = num(args[0]);
    return isErr(n) ? n : ymd(n).y;
  },
  MONTH: (args) => {
    const n = num(args[0]);
    return isErr(n) ? n : ymd(n).m;
  },
  DAY: (args) => {
    const n = num(args[0]);
    return isErr(n) ? n : ymd(n).d;
  },
  WEEKDAY: (args) => {
    const n = num(args[0]);
    const type = num(args[1], 1);
    if (isErr(n)) return n;
    if (isErr(type)) return type;
    const js = serialToDate(n).getUTCDay(); // 0 = вс
    if (type === 2) return js === 0 ? 7 : js;
    if (type === 3) return js === 0 ? 6 : js - 1;
    return js + 1;
  },
  EDATE: (args) => {
    const n = num(args[0]);
    const k = num(args[1]);
    if (isErr(n)) return n;
    if (isErr(k)) return k;
    const { y, m, d } = ymd(n);
    const last = new Date(Date.UTC(y, m - 1 + k + 1, 0)).getUTCDate();
    return makeSerial(y, m + k, Math.min(d, last));
  },
  DAYS: (args) => {
    const a = num(args[0]);
    const b = num(args[1]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return Math.floor(a) - Math.floor(b);
  },
  DATEDIF: (args) => {
    const a = num(args[0]);
    const b = num(args[1]);
    const unit = txt(args[2]);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    if (isErr(unit)) return unit;
    if (b < a) return ERR.NUM;
    const A = ymd(a);
    const B = ymd(b);
    switch (unit.toUpperCase()) {
      case 'D':
        return Math.floor(b) - Math.floor(a);
      case 'M':
        return (B.y - A.y) * 12 + (B.m - A.m) - (B.d < A.d ? 1 : 0);
      case 'Y':
        return B.y - A.y - (B.m < A.m || (B.m === A.m && B.d < A.d) ? 1 : 0);
      default:
        return ERR.NUM;
    }
  },
};

FUNCTIONS.CONCAT = FUNCTIONS.CONCATENATE;

export const LAZY_FUNCTIONS: Record<string, LazyImpl> = {
  IF: (nodes, _ctx, ev) => {
    if (nodes.length < 1) return ERR.VALUE;
    const cond = toBool(scalarOf(ev(nodes[0])));
    if (isErr(cond)) return cond;
    if (cond) return nodes[1] ? ev(nodes[1]) : true;
    return nodes[2] ? ev(nodes[2]) : false;
  },
  IFS: (nodes, _ctx, ev) => {
    for (let i = 0; i + 1 < nodes.length; i += 2) {
      const cond = toBool(scalarOf(ev(nodes[i])));
      if (isErr(cond)) return cond;
      if (cond) return ev(nodes[i + 1]);
    }
    return ERR.NA;
  },
  IFERROR: (nodes, _ctx, ev) => {
    const v = scalarOf(ev(nodes[0]));
    return isErr(v) ? (nodes[1] ? ev(nodes[1]) : '') : v;
  },
  IFNA: (nodes, _ctx, ev) => {
    const v = scalarOf(ev(nodes[0]));
    return isErr(v) && v.code === 'NA' ? (nodes[1] ? ev(nodes[1]) : '') : v;
  },
  SWITCH: (nodes, _ctx, ev) => {
    const v = scalarOf(ev(nodes[0]));
    if (isErr(v)) return v;
    let i = 1;
    for (; i + 1 < nodes.length; i += 2) {
      if (compareScalars(v, scalarOf(ev(nodes[i]))) === 0) return ev(nodes[i + 1]);
    }
    return i < nodes.length ? ev(nodes[i]) : ERR.NA;
  },
  CHOOSE: (nodes, _ctx, ev) => {
    const k = toNumber(scalarOf(ev(nodes[0])));
    if (isErr(k)) return k;
    const idx = Math.trunc(k);
    if (idx < 1 || idx >= nodes.length) return ERR.VALUE;
    return ev(nodes[idx]);
  },
  ROW: (nodes, ctx) => {
    const n = nodes[0];
    if (!n) return ctx.row + 1;
    if (n.t !== 'ref') return ERR.VALUE;
    if (n.a.r === undefined) return ERR.VALUE;
    return (n.a.ra ? n.a.r : n.a.r + ctx.baseRow) + 1;
  },
  COLUMN: (nodes, ctx) => {
    const n = nodes[0];
    if (!n) return ctx.col + 1;
    if (n.t !== 'ref' || n.a.c === undefined) return ERR.VALUE;
    return n.a.c + 1;
  },
};

/** Русские имена функций, как в русском Excel. */
export const RU_NAMES: Record<string, string> = {
  СУММ: 'SUM',
  СРЗНАЧ: 'AVERAGE',
  МИН: 'MIN',
  МАКС: 'MAX',
  ПРОИЗВЕД: 'PRODUCT',
  МЕДИАНА: 'MEDIAN',
  СЧЁТ: 'COUNT',
  СЧЕТ: 'COUNT',
  СЧЁТЗ: 'COUNTA',
  СЧЕТЗ: 'COUNTA',
  СЧИТАТЬПУСТОТЫ: 'COUNTBLANK',
  ОКРУГЛ: 'ROUND',
  ОКРУГЛВВЕРХ: 'ROUNDUP',
  ОКРУГЛВНИЗ: 'ROUNDDOWN',
  ОКРВВЕРХ: 'CEILING',
  ОКРВНИЗ: 'FLOOR',
  ЦЕЛОЕ: 'INT',
  ОСТАТ: 'MOD',
  СТЕПЕНЬ: 'POWER',
  КОРЕНЬ: 'SQRT',
  СУММПРОИЗВ: 'SUMPRODUCT',
  СУММЕСЛИ: 'SUMIF',
  СУММЕСЛИМН: 'SUMIFS',
  СРЗНАЧЕСЛИ: 'AVERAGEIF',
  СРЗНАЧЕСЛИМН: 'AVERAGEIFS',
  МАКСЕСЛИ: 'MAXIFS',
  МИНЕСЛИ: 'MINIFS',
  СЧЁТЕСЛИ: 'COUNTIF',
  СЧЕТЕСЛИ: 'COUNTIF',
  СЧЁТЕСЛИМН: 'COUNTIFS',
  СЧЕТЕСЛИМН: 'COUNTIFS',
  'ПРОМЕЖУТОЧНЫЕ.ИТОГИ': 'SUBTOTAL',
  ЕСЛИ: 'IF',
  ЕСЛИМН: 'IFS',
  ЕСЛИОШИБКА: 'IFERROR',
  ЕСНД: 'IFNA',
  ПЕРЕКЛЮЧ: 'SWITCH',
  ВЫБОР: 'CHOOSE',
  И: 'AND',
  ИЛИ: 'OR',
  НЕ: 'NOT',
  ИСТИНА: 'TRUE',
  ЛОЖЬ: 'FALSE',
  ЕПУСТО: 'ISBLANK',
  ЕЧИСЛО: 'ISNUMBER',
  ЕТЕКСТ: 'ISTEXT',
  ЕОШИБКА: 'ISERROR',
  НД: 'NA',
  СЦЕП: 'CONCAT',
  СЦЕПИТЬ: 'CONCATENATE',
  ОБЪЕДИНИТЬ: 'TEXTJOIN',
  ЛЕВСИМВ: 'LEFT',
  ПРАВСИМВ: 'RIGHT',
  ПСТР: 'MID',
  ДЛСТР: 'LEN',
  ПРОПИСН: 'UPPER',
  СТРОЧН: 'LOWER',
  ПРОПНАЧ: 'PROPER',
  СЖПРОБЕЛЫ: 'TRIM',
  ПОДСТАВИТЬ: 'SUBSTITUTE',
  ЗАМЕНИТЬ: 'REPLACE',
  НАЙТИ: 'FIND',
  ПОИСК: 'SEARCH',
  ПОВТОР: 'REPT',
  СОВПАД: 'EXACT',
  ЗНАЧЕН: 'VALUE',
  ТЕКСТ: 'TEXT',
  ВПР: 'VLOOKUP',
  ГПР: 'HLOOKUP',
  ПОИСКПОЗ: 'MATCH',
  ИНДЕКС: 'INDEX',
  ПРОСМОТРX: 'XLOOKUP',
  ПРОСМОТРХ: 'XLOOKUP',
  СТРОКА: 'ROW',
  СТОЛБЕЦ: 'COLUMN',
  ЧСТРОК: 'ROWS',
  ЧИСЛСТОЛБ: 'COLUMNS',
  СЕГОДНЯ: 'TODAY',
  ТДАТА: 'NOW',
  ДАТА: 'DATE',
  ГОД: 'YEAR',
  МЕСЯЦ: 'MONTH',
  ДЕНЬ: 'DAY',
  ДЕНЬНЕД: 'WEEKDAY',
  ДАТАМЕС: 'EDATE',
  ДНИ: 'DAYS',
  РАЗНДАТ: 'DATEDIF',
};

export function resolveFunctionName(name: string): string | null {
  const up = name.toUpperCase().replace(/^_XLFN\./, '');
  if (FUNCTIONS[up] || LAZY_FUNCTIONS[up]) return up;
  const ru = RU_NAMES[up];
  return ru ?? null;
}

/** Список для подсказок в строке формул: [русское имя, английское, краткое описание]. */
export const FUNCTION_HINTS: [string, string, string][] = [
  ['СУММ', 'SUM', 'Сумма чисел или диапазона'],
  ['СРЗНАЧ', 'AVERAGE', 'Среднее значение'],
  ['МИН', 'MIN', 'Наименьшее значение'],
  ['МАКС', 'MAX', 'Наибольшее значение'],
  ['СЧЁТ', 'COUNT', 'Количество чисел'],
  ['СЧЁТЗ', 'COUNTA', 'Количество непустых ячеек'],
  ['ЕСЛИ', 'IF', 'ЕСЛИ(условие; если да; если нет)'],
  ['ЕСЛИОШИБКА', 'IFERROR', 'Значение или замена при ошибке'],
  ['ОКРУГЛ', 'ROUND', 'ОКРУГЛ(число; знаков)'],
  ['ОКРУГЛВВЕРХ', 'ROUNDUP', 'Округление вверх'],
  ['ОКРУГЛВНИЗ', 'ROUNDDOWN', 'Округление вниз'],
  ['СУММЕСЛИ', 'SUMIF', 'СУММЕСЛИ(диапазон; условие; [сумм. диапазон])'],
  ['СЧЁТЕСЛИ', 'COUNTIF', 'СЧЁТЕСЛИ(диапазон; условие)'],
  ['СУММЕСЛИМН', 'SUMIFS', 'Сумма по нескольким условиям'],
  ['СЧЁТЕСЛИМН', 'COUNTIFS', 'Количество по нескольким условиям'],
  ['СРЗНАЧЕСЛИ', 'AVERAGEIF', 'Среднее по условию'],
  ['ВПР', 'VLOOKUP', 'ВПР(что; таблица; номер столбца; 0)'],
  ['ПРОСМОТРX', 'XLOOKUP', 'ПРОСМОТРX(что; где; откуда вернуть)'],
  ['ИНДЕКС', 'INDEX', 'ИНДЕКС(диапазон; строка; [столбец])'],
  ['ПОИСКПОЗ', 'MATCH', 'ПОИСКПОЗ(что; где; 0)'],
  ['И', 'AND', 'Все условия истинны'],
  ['ИЛИ', 'OR', 'Хотя бы одно условие истинно'],
  ['НЕ', 'NOT', 'Отрицание'],
  ['СЦЕП', 'CONCAT', 'Склеить текст'],
  ['ОБЪЕДИНИТЬ', 'TEXTJOIN', 'Склеить через разделитель'],
  ['ЛЕВСИМВ', 'LEFT', 'Первые символы'],
  ['ПРАВСИМВ', 'RIGHT', 'Последние символы'],
  ['ПСТР', 'MID', 'Часть текста'],
  ['ДЛСТР', 'LEN', 'Длина текста'],
  ['ПРОПИСН', 'UPPER', 'ВЕРХНИЙ регистр'],
  ['СТРОЧН', 'LOWER', 'нижний регистр'],
  ['СЖПРОБЕЛЫ', 'TRIM', 'Убрать лишние пробелы'],
  ['ПОДСТАВИТЬ', 'SUBSTITUTE', 'Заменить текст'],
  ['ТЕКСТ', 'TEXT', 'Число в текст по шаблону'],
  ['ЗНАЧЕН', 'VALUE', 'Текст в число'],
  ['СЕГОДНЯ', 'TODAY', 'Сегодняшняя дата'],
  ['ДАТА', 'DATE', 'ДАТА(год; месяц; день)'],
  ['ГОД', 'YEAR', 'Год из даты'],
  ['МЕСЯЦ', 'MONTH', 'Месяц из даты'],
  ['ДЕНЬ', 'DAY', 'День из даты'],
  ['РАЗНДАТ', 'DATEDIF', 'Разница дат: "D", "M", "Y"'],
  ['ABS', 'ABS', 'Модуль числа'],
  ['ЦЕЛОЕ', 'INT', 'Целая часть'],
  ['ОСТАТ', 'MOD', 'Остаток от деления'],
  ['СТЕПЕНЬ', 'POWER', 'Возведение в степень'],
  ['КОРЕНЬ', 'SQRT', 'Квадратный корень'],
  ['СУММПРОИЗВ', 'SUMPRODUCT', 'Сумма произведений'],
  ['ПРОМЕЖУТОЧНЫЕ.ИТОГИ', 'SUBTOTAL', 'Итоги только по видимым строкам (9 — сумма)'],
  ['МЕДИАНА', 'MEDIAN', 'Медиана'],
  ['СТРОКА', 'ROW', 'Номер строки'],
];
