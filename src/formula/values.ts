export type ErrCode = 'DIV0' | 'VALUE' | 'REF' | 'NAME' | 'NA' | 'NUM' | 'NULL' | 'CYCLE';

export class FErr {
  constructor(public readonly code: ErrCode) {}
  toString(): string {
    return ERROR_TEXT[this.code];
  }
}

/** Как ошибки выглядят в русском Excel — так их и показываем. */
export const ERROR_TEXT: Record<ErrCode, string> = {
  DIV0: '#ДЕЛ/0!',
  VALUE: '#ЗНАЧ!',
  REF: '#ССЫЛКА!',
  NAME: '#ИМЯ?',
  NA: '#Н/Д',
  NUM: '#ЧИСЛО!',
  NULL: '#ПУСТО!',
  CYCLE: '#ЦИКЛ!',
};

/** Английская запись для экспорта в .xlsx. */
export const ERROR_TEXT_EN: Record<ErrCode, string> = {
  DIV0: '#DIV/0!',
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  NA: '#N/A',
  NUM: '#NUM!',
  NULL: '#NULL!',
  CYCLE: '#REF!',
};

export const ERROR_LITERALS: [string, ErrCode][] = (
  [
    ...Object.entries(ERROR_TEXT),
    ...Object.entries(ERROR_TEXT_EN).filter(([k]) => k !== 'CYCLE'),
  ] as [ErrCode, string][]
)
  .map(([code, text]) => [text.toUpperCase(), code] as [string, ErrCode])
  .sort((a, b) => b[0].length - a[0].length);

export const ERR = {
  DIV0: new FErr('DIV0'),
  VALUE: new FErr('VALUE'),
  REF: new FErr('REF'),
  NAME: new FErr('NAME'),
  NA: new FErr('NA'),
  NUM: new FErr('NUM'),
  NULL: new FErr('NULL'),
  CYCLE: new FErr('CYCLE'),
} as const;

/** Скалярное значение ячейки. null — пустая ячейка. */
export type Scalar = number | string | boolean | FErr | null;

/** Прямоугольный диапазон, вычисляемый лениво. */
export interface RangeVal {
  kind: 'range';
  rows: number;
  cols: number;
  get(r: number, c: number): Scalar;
  /** Абсолютные координаты верхнего левого угла — для СТРОКА()/СТОЛБЕЦ(). */
  top: number;
  left: number;
  sheetId: string;
}

export type Value = Scalar | RangeVal;

export function isRange(v: Value): v is RangeVal {
  return typeof v === 'object' && v !== null && (v as RangeVal).kind === 'range';
}

export function isErr(v: unknown): v is FErr {
  return v instanceof FErr;
}
