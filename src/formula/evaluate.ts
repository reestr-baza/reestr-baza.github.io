import { compareScalars, scalarOf, toNumber, toText } from './coerce';
import { FUNCTIONS, LAZY_FUNCTIONS, resolveFunctionName } from './functions';
import type { Node } from './parse';
import { makeRange } from './range';
import { ERR, FErr, isErr, type RangeVal, type Scalar, type Value } from './values';

/** Всё, что вычислителю нужно знать о книге. Реализуется хранилищем. */
export interface EvalHost {
  /** id листа по имени (без учёта регистра); undefined → текущий лист. */
  sheetByName(name: string): string | null;
  cellValue(sheetId: string, row: number, col: number): Scalar;
  rowCount(sheetId: string): number;
  colCount(sheetId: string): number;
  namedValue(name: string): Scalar | undefined;
  /** Скрыта ли строка: вручную или фильтром — для ПРОМЕЖУТОЧНЫЕ.ИТОГИ. */
  rowHidden(sheetId: string, row: number): { manual: boolean; filtered: boolean };
  /** Индекс для точного поиска в столбце диапазона (ВПР, ПОИСКПОЗ). */
  lookupIndex?(range: RangeVal, col: number): Map<string, number>;
}

export interface EvalCtx {
  host: EvalHost;
  sheetId: string;
  /** Строка/столбец ячейки с формулой — для СТРОКА() и относительных ссылок. */
  row: number;
  col: number;
  /** Сдвиг относительных строк: для формулы столбца, записанной относительно первой строки. */
  baseRow: number;
}

function resolveSheet(ctx: EvalCtx, name?: string): string | FErr {
  if (!name) return ctx.sheetId;
  return ctx.host.sheetByName(name) ?? ERR.REF;
}

function evalRef(node: Extract<Node, { t: 'ref' }>, ctx: EvalCtx): Value {
  const sheetId = resolveSheet(ctx, node.sheet);
  if (isErr(sheetId)) return sheetId;
  const { a, b } = node;
  const rowOf = (r: number | undefined, ra: boolean) => (r === undefined ? undefined : ra ? r : r + ctx.baseRow);

  if (!b) {
    const r = rowOf(a.r, a.ra)!;
    if (r < 0 || a.c! < 0) return ERR.REF;
    return ctx.host.cellValue(sheetId, r, a.c!);
  }
  const lastRow = Math.max(0, ctx.host.rowCount(sheetId) - 1);
  const lastCol = Math.max(0, ctx.host.colCount(sheetId) - 1);
  let r1 = rowOf(a.r, a.ra);
  let r2 = rowOf(b.r, b.ra);
  let c1 = a.c;
  let c2 = b.c;
  if (r1 === undefined || r2 === undefined) {
    // столбцы целиком: A:C — берём только заполненную часть листа
    r1 = 0;
    r2 = lastRow;
  }
  if (c1 === undefined || c2 === undefined) {
    c1 = 0;
    c2 = lastCol;
  }
  if (r1 < 0 || r2 < 0) return ERR.REF;
  return makeRange(ctx, sheetId, r1, c1, r2, c2);
}

function arith(op: string, x: Scalar, y: Scalar): Scalar {
  const a = toNumber(x);
  if (isErr(a)) return a;
  const b = toNumber(y);
  if (isErr(b)) return b;
  let r: number;
  switch (op) {
    case '+':
      r = a + b;
      break;
    case '-':
      r = a - b;
      break;
    case '*':
      r = a * b;
      break;
    case '/':
      if (b === 0) return ERR.DIV0;
      r = a / b;
      break;
    case '^':
      r = Math.pow(a, b);
      break;
    default:
      return ERR.VALUE;
  }
  if (!Number.isFinite(r)) return ERR.NUM;
  // убираем хвосты двоичной арифметики: 0,1+0,2 = 0,3
  return Math.abs(r) < 1e15 ? parseFloat(r.toPrecision(15)) : r;
}

export function evaluate(node: Node, ctx: EvalCtx): Value {
  switch (node.t) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return new FErr(node.v);
    case 'ref':
      return evalRef(node, ctx);
    case 'name': {
      const v = ctx.host.namedValue(node.name);
      return v === undefined ? ERR.NAME : v;
    }
    case 'neg': {
      const v = toNumber(scalarOf(evaluate(node.a, ctx)));
      return isErr(v) ? v : v === 0 ? 0 : -v;
    }
    case 'pct': {
      const v = toNumber(scalarOf(evaluate(node.a, ctx)));
      return isErr(v) ? v : v / 100;
    }
    case 'bin': {
      const x = scalarOf(evaluate(node.a, ctx));
      const y = scalarOf(evaluate(node.b, ctx));
      if (isErr(x)) return x;
      if (isErr(y)) return y;
      switch (node.op) {
        case '&': {
          const sx = toText(x);
          const sy = toText(y);
          if (isErr(sx)) return sx;
          if (isErr(sy)) return sy;
          return sx + sy;
        }
        case '=':
          return compareScalars(x, y) === 0;
        case '<>':
          return compareScalars(x, y) !== 0;
        case '<':
          return compareScalars(x, y) < 0;
        case '>':
          return compareScalars(x, y) > 0;
        case '<=':
          return compareScalars(x, y) <= 0;
        case '>=':
          return compareScalars(x, y) >= 0;
        default:
          return arith(node.op, x, y);
      }
    }
    case 'fn': {
      const name = resolveFunctionName(node.name);
      if (!name) return ERR.NAME;
      const lazy = LAZY_FUNCTIONS[name];
      if (lazy) return lazy(node.args, ctx, (n) => evaluate(n, ctx));
      const impl = FUNCTIONS[name];
      const args = node.args.map((a) => evaluate(a, ctx));
      return impl(args, ctx);
    }
  }
}

/** Итоговое значение ячейки: диапазон в ячейке показываем как его левый верхний элемент. */
export function evaluateToScalar(node: Node, ctx: EvalCtx): Scalar {
  const v = evaluate(node, ctx);
  if (typeof v === 'object' && v !== null && !(v instanceof FErr)) {
    return v.rows >= 1 && v.cols >= 1 ? v.get(0, 0) : ERR.VALUE;
  }
  return v;
}
