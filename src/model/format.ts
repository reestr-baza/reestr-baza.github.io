import { numberToText, parseLooseNumber } from '../formula/coerce';
import { dateToSerial, serialToDate } from '../formula/functions';
import { isErr, type Scalar } from '../formula/values';
import type { Currency, NumFmt } from './types';

const NBSP = ' ';
const MINUS = '−';

export const CURRENCIES: { c: Currency; symbol: string; label: string; prefix: boolean }[] = [
  { c: 'RUB', symbol: '₽', label: 'Рубли', prefix: false },
  { c: 'CNY', symbol: '¥', label: 'Юани', prefix: true },
  { c: 'USD', symbol: '$', label: 'Доллары', prefix: true },
  { c: 'EUR', symbol: '€', label: 'Евро', prefix: false },
];

export function currencyInfo(c: Currency) {
  return CURRENCIES.find((x) => x.c === c)!;
}

function grouped(n: number, decimals: number): string {
  const fixed = Math.abs(n).toFixed(decimals);
  const [ip, fp] = fixed.split('.');
  const intText = ip.replace(/\B(?=(\d{3})+(?!\d))/g, NBSP);
  return fp ? `${intText},${fp}` : intText;
}

export function formatDate(serial: number): string {
  const d = serialToDate(serial);
  const pad = (x: number) => String(x).padStart(2, '0');
  return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
}

export function formatNumber(n: number, nf?: NumFmt): string {
  if (!Number.isFinite(n)) return '#ЧИСЛО!';
  const sign = n < 0 ? MINUS : '';
  switch (nf?.k) {
    case 'number':
      return sign + grouped(n, nf.d);
    case 'currency': {
      const info = currencyInfo(nf.c);
      const body = grouped(n, nf.d);
      return info.prefix ? `${sign}${info.symbol}${body}` : `${sign}${body}${NBSP}${info.symbol}`;
    }
    case 'percent':
      return `${sign}${grouped(n * 100, nf.d)}%`;
    case 'date':
      return formatDate(n);
    default: {
      const t = numberToText(n);
      return t.startsWith('-') ? MINUS + t.slice(1) : t;
    }
  }
}

export function formatScalar(v: Scalar, nf?: NumFmt): string {
  if (v === null) return '';
  if (typeof v === 'number') return nf?.k === 'text' ? numberToText(v) : formatNumber(v, nf);
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (isErr(v)) return v.toString();
  return v;
}

/** Текст для редактирования: число без форматирования, с десятичной запятой. */
export function editText(v: string | number | boolean | undefined, nf?: NumFmt): string {
  if (v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'number') {
    if (nf?.k === 'date') return formatDate(v);
    if (nf?.k === 'percent') return numberToText(parseFloat((v * 100).toPrecision(12))) + '%';
    return numberToText(v);
  }
  return v;
}

export interface ParsedInput {
  v?: string | number | boolean;
  f?: string;
  /** Формат, угаданный из ввода (¥2 500 → юани). Применяется, если у ячейки формат «общий». */
  nf?: NumFmt;
}

const CUR_SUFFIX: [RegExp, Currency][] = [
  [/^(₽|р\.?|руб\.?|rub)$/i, 'RUB'],
  [/^(¥|юан[а-я]*|cny|rmb)$/i, 'CNY'],
  [/^(\$|usd)$/i, 'USD'],
  [/^(€|eur)$/i, 'EUR'],
];

function currencyBySymbol(s: string): Currency | null {
  for (const [re, c] of CUR_SUFFIX) if (re.test(s.trim())) return c;
  return null;
}

export const URL_RE = /^(https?:\/\/|www\.)[^\s]+$/i;

/** Разбор введённого пользователем текста в значение ячейки. */
export function parseInput(raw: string): ParsedInput {
  const text = raw.replace(/\r\n/g, '\n');
  if (text === '') return {};
  if (text.startsWith("'")) return { v: text.slice(1) };
  if (text.startsWith('=') && text.length > 1) return { f: text };
  const t = text.trim();

  const up = t.toUpperCase();
  if (up === 'ИСТИНА' || up === 'TRUE') return { v: true };
  if (up === 'ЛОЖЬ' || up === 'FALSE') return { v: false };

  // Дата 18.09.2026 / 18.09.26
  const dm = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/.exec(t);
  if (dm) {
    const d = +dm[1];
    const m = +dm[2];
    let y = +dm[3];
    if (y < 100) y += 2000;
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return { v: dateToSerial(new Date(y, m - 1, d)), nf: { k: 'date' } };
    }
  }

  // Процент
  if (/%$/.test(t)) {
    const n = parseLooseNumber(t.slice(0, -1).replace(/^[−–]/, '-'));
    if (n !== null) {
      const decimals = (/[.,](\d+)%$/.exec(t)?.[1].length ?? 0);
      return { v: n / 100, nf: { k: 'percent', d: decimals } };
    }
  }

  // Число с валютой: ¥2 500, 31 875 ₽, $12.5, 100 руб.
  const cm = /^([−–-]?)\s*([¥$€₽]?)\s*([\d\s .,]+?)\s*([^\d\s.,]*)$/.exec(t);
  if (cm && /\d/.test(cm[3])) {
    const n = parseLooseNumber(cm[3]);
    if (n !== null) {
      const sym = cm[2] || cm[4];
      const value = cm[1] ? -n : n;
      if (!sym) return { v: value };
      const c = currencyBySymbol(sym);
      if (c) {
        const decimals = /[.,]\d{1,2}$/.test(cm[3].trim()) ? 2 : 0;
        return { v: value, nf: { k: 'currency', c, d: decimals } };
      }
    }
  }
  return { v: text };
}

export function isUrl(s: string): boolean {
  return URL_RE.test(s.trim());
}

export function normalizeUrl(s: string): string {
  const t = s.trim();
  if (/^www\./i.test(t)) return 'https://' + t;
  if (!/^[a-z][a-z0-9+.-]*:/i.test(t)) return 'https://' + t;
  return t;
}

/** Ссылка безопасна для открытия: только http(s) и mailto. */
export function safeHref(s: string): string | null {
  try {
    const u = new URL(normalizeUrl(s));
    return u.protocol === 'http:' || u.protocol === 'https:' || u.protocol === 'mailto:' ? u.href : null;
  } catch {
    return null;
  }
}
