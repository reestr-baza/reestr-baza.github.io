import { resolveFunctionName, RU_NAMES } from './functions';
import { adjustFormula, refText } from './refs';
import { formulaTokens } from './parse';
import { tokenize, type Tok } from './tokenize';
import { ERROR_TEXT, ERROR_TEXT_EN } from './values';

/**
 * Формула «Реестра» → формула для файла .xlsx: английские имена функций,
 * запятая между аргументами, точка в числах. `rowShift` — сдвиг строк (строка заголовков в Excel).
 */
export function toExcelFormula(src: string, rowShift: number): string {
  let f = src;
  if (rowShift) f = adjustFormula(f, { axis: 'row', at: 0, count: rowShift }, () => true);
  let toks: Tok[];
  try {
    toks = formulaTokens(f);
  } catch {
    try {
      toks = tokenize(f, true);
    } catch {
      return f.replace(/^=/, '');
    }
  }
  let out = '';
  for (const t of toks) {
    switch (t.k) {
      case 'num':
        out += String(t.v);
        break;
      case 'str':
        out += `"${t.v.replace(/"/g, '""')}"`;
        break;
      case 'bool':
        out += t.v ? 'TRUE' : 'FALSE';
        break;
      case 'err':
        out += ERROR_TEXT_EN[t.v];
        break;
      case 'ref':
        out += refText({ sheet: t.sheet, a: t.a, b: t.b });
        break;
      case 'fn': {
        const en = resolveFunctionName(t.name) ?? t.name;
        // новые функции Excel хранятся с префиксом _xlfn.
        out += ['XLOOKUP', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCAT', 'IFNA', 'DAYS'].includes(en) ? `_xlfn.${en}` : en;
        break;
      }
      case 'name':
        out += t.name;
        break;
      case 'op':
        out += t.v;
        break;
      case '(':
        out += '(';
        break;
      case ')':
        out += ')';
        break;
      case 'sep':
        out += ',';
        break;
    }
  }
  return out;
}

/** Формула из .xlsx → формула «Реестра» (сдвиг строк, если первые строки стали заголовками). */
export function fromExcelFormula(src: string, rowShift: number): string {
  let f = '=' + src.replace(/_xlfn\./gi, '').replace(/_xlws\./gi, '');
  if (rowShift) f = adjustFormula(f, { axis: 'row', at: 0, count: rowShift }, () => true);
  return toRuFormula(f);
}

let enToRu: Map<string, string> | null = null;

/** Английская запись формулы → как в русском Excel: ЕСЛИ, «;», десятичная запятая. */
export function toRuFormula(src: string): string {
  if (!enToRu) {
    enToRu = new Map();
    for (const [ru, en] of Object.entries(RU_NAMES)) if (!enToRu.has(en)) enToRu.set(en, ru);
  }
  let toks: Tok[];
  try {
    toks = formulaTokens(src);
  } catch {
    return src;
  }
  let out = '=';
  let pos = src.startsWith('=') ? 1 : 0;
  for (const t of toks) {
    // пробелы между токенами сохраняем как были
    out += src.slice(pos, t.s);
    pos = t.e;
    switch (t.k) {
      case 'num':
        out += String(t.v).replace('.', ',');
        break;
      case 'str':
        out += `"${t.v.replace(/"/g, '""')}"`;
        break;
      case 'bool':
        out += t.v ? 'ИСТИНА' : 'ЛОЖЬ';
        break;
      case 'err':
        out += ERROR_TEXT[t.v];
        break;
      case 'ref':
        out += refText({ sheet: t.sheet, a: t.a, b: t.b });
        break;
      case 'fn': {
        const en = resolveFunctionName(t.name);
        out += (en && enToRu.get(en)) ?? t.name;
        break;
      }
      case 'name':
        out += t.name;
        break;
      case 'op':
        out += t.v;
        break;
      case '(':
      case ')':
        out += t.k;
        break;
      case 'sep':
        out += ';';
        break;
    }
  }
  return out + src.slice(pos);
}
