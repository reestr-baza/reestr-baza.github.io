export type ColId = string;
export type RowId = string;
export type SheetId = string;
export type ImageId = string;

export type HAlign = 'left' | 'center' | 'right';
export type VAlign = 'top' | 'middle' | 'bottom';
export type Currency = 'RUB' | 'CNY' | 'USD' | 'EUR';

export type NumFmt =
  | { k: 'general' }
  | { k: 'number'; d: number }
  | { k: 'currency'; c: Currency; d: number }
  | { k: 'percent'; d: number }
  | { k: 'date' }
  | { k: 'text' };

export interface CellStyle {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  /** Цвет шрифта */
  fg?: string;
  /** Заливка */
  bg?: string;
  ha?: HAlign;
  va?: VAlign;
  wrap?: boolean;
  nf?: NumFmt;
}

export interface Cell {
  /** Введённое значение (не формула). */
  v?: string | number | boolean;
  /** Формула, начиная с «=». */
  f?: string;
  img?: ImageId;
  /** Явная ссылка: текст ячейки — подпись, href — адрес. */
  href?: string;
  /** Примечание к ячейке (как в Excel) */
  note?: string;
  st?: CellStyle;
}

/** Блок карточки товара: фото + текст с выравниванием, как ячейка. */
export interface CardBlock {
  img?: ImageId;
  text?: string;
  st?: CellStyle;
}

export interface Row {
  id: RowId;
  cells: Record<ColId, Cell>;
  /** Своя высота строки, px */
  h?: number;
  hidden?: boolean;
  card?: CardBlock[];
}

export interface Column {
  id: ColId;
  name: string;
  /** Ширина, px */
  w: number;
  hidden?: boolean;
  /** Оформление по умолчанию для всех ячеек столбца */
  st?: CellStyle;
  /** Формула столбца, записанная для первой строки (строка 1). Применяется к пустым ячейкам. */
  formula?: string;
}

export type FilterOp =
  | 'contains'
  | 'notContains'
  | 'eq'
  | 'neq'
  | 'begins'
  | 'ends'
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'
  | 'between'
  | 'empty'
  | 'notEmpty'
  | 'top';

export interface FilterSpec {
  /** Разрешённые значения (ключи из valueKey). Отсутствует — фильтра по значениям нет. */
  values?: string[];
  cond?: { op: FilterOp; a?: string; b?: string };
  color?: { kind: 'bg' | 'fg'; color: string | null };
  photo?: 'with' | 'without';
}

export type Density = 'S' | 'M' | 'L';

export interface SheetMeta {
  id: SheetId;
  name: string;
  columns: Column[];
  rowOrder: RowId[];
  /** Столбец, по клику на значение которого открывается карточка */
  keyColId?: ColId;
  /** Сколько видимых столбцов слева закреплено */
  frozen: number;
  density: Density;
  filters: Record<ColId, FilterSpec>;
}

export interface Sheet extends SheetMeta {
  rows: Map<RowId, Row>;
}

export interface NamedValue {
  name: string;
  value: number | string;
  note?: string;
}

export interface WorkbookMeta {
  id: string;
  title: string;
  sheetIds: SheetId[];
  activeSheet: SheetId;
  names: NamedValue[];
  createdAt: number;
  /** Книга — стартовый пример: импорт по умолчанию заменит его */
  demo?: boolean;
}

export const DENSITY_HEIGHT: Record<Density, number> = { S: 30, M: 52, L: 104 };
export const DEFAULT_COL_WIDTH = 120;
export const MAX_COLUMNS = 200;
