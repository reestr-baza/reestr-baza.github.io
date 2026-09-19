import { ERR, FErr, isErr, isRange, type Scalar, type Value } from './values';

/** Разбор числа из текста с учётом русской записи: «31 875», «12,75», «1 234,5», «15%». */
export function parseLooseNumber(s: string): number | null {
  const spaced = s.trim().replace(/[  ]/g, ' ');
  // пробелы допустимы только как разделители тысяч: «31 875», но не «42 44 46»
  if (/\s/.test(spaced) && !/^[+-]?\d{1,3}( \d{3})+([.,]\d+)?%?$/.test(spaced)) return null;
  const t = spaced.replace(/\s/g, '');
  if (t === '') return null;
  let m = /^([+-]?)(\d+)(?:[.,](\d+))?(?:[eE]([+-]?\d+))?(%?)$/.exec(t);
  if (m) {
    let v = parseFloat(`${m[1]}${m[2]}.${m[3] ?? '0'}${m[4] ? 'e' + m[4] : ''}`);
    if (m[5]) v /= 100;
    return v;
  }
  // «1,234.56» — американская запись с разделителями тысяч
  m = /^([+-]?)(\d{1,3}(?:,\d{3})+)(?:\.(\d+))?$/.exec(t);
  if (m) return parseFloat(`${m[1]}${m[2].replace(/,/g, '')}.${m[3] ?? '0'}`);
  return null;
}

export function toNumber(v: Scalar): number | FErr {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v)) return v;
  const n = parseLooseNumber(v);
  return n === null ? ERR.VALUE : n;
}

/** Число → строка «как в общем формате»: до 10 значащих цифр, десятичная запятая. */
export function numberToText(n: number): string {
  if (!Number.isFinite(n)) return '#ЧИСЛО!';
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let s = Math.abs(n) >= 1e-5 && Math.abs(n) < 1e15 ? String(parseFloat(n.toPrecision(10))) : n.toExponential(5);
  s = s.replace('.', ',');
  return s;
}

export function toText(v: Scalar): string | FErr {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numberToText(v);
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  return v;
}

export function toBool(v: Scalar): boolean | FErr {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) return v;
  const u = v.trim().toUpperCase();
  if (u === 'TRUE' || u === 'ИСТИНА') return true;
  if (u === 'FALSE' || u === 'ЛОЖЬ') return false;
  return ERR.VALUE;
}

/** Диапазон 1×1 в скалярном контексте превращается в значение; больший — в #ЗНАЧ!. */
export function scalarOf(v: Value): Scalar {
  if (!isRange(v)) return v;
  if (v.rows === 1 && v.cols === 1) return v.get(0, 0);
  return ERR.VALUE;
}

/** Порядок типов при сравнении в Excel: числа < текст < логические. */
function typeRank(v: Scalar): number {
  if (v === null) return 0;
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 4;
}

const collator = new Intl.Collator('ru', { sensitivity: 'base', numeric: true });

export function compareScalars(a: Scalar, b: Scalar): number {
  // пустая ячейка сравнивается как 0 с числом и как "" с текстом
  if (a === null && typeof b === 'number') a = 0;
  if (b === null && typeof a === 'number') b = 0;
  if (a === null && typeof b === 'string') a = '';
  if (b === null && typeof a === 'string') b = '';
  if (a === null && typeof b === 'boolean') a = false;
  if (b === null && typeof a === 'boolean') b = false;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return collator.compare(a, b);
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return 0;
}

export function textEquals(a: string, b: string): boolean {
  return collator.compare(a, b) === 0;
}
