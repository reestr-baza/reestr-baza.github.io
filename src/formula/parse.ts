import { FormulaSyntaxError, tokenize, usesSemicolons, type RefEnd, type Tok } from './tokenize';
import type { ErrCode } from './values';

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrCode }
  | { t: 'ref'; sheet?: string; a: RefEnd; b?: RefEnd }
  | { t: 'name'; name: string }
  | { t: 'fn'; name: string; args: Node[] }
  | { t: 'neg'; a: Node }
  | { t: 'pct'; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node };

const BIN_PREC: Record<string, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
};

class Parser {
  private i = 0;
  constructor(private toks: Tok[]) {}

  parse(): Node {
    if (this.toks.length === 0) throw new FormulaSyntaxError('Пустая формула');
    const node = this.expr(0);
    if (this.i < this.toks.length) throw new FormulaSyntaxError('Лишние символы в формуле');
    return node;
  }

  private peek(): Tok | undefined {
    return this.toks[this.i];
  }

  private expr(minPrec: number): Node {
    let left = this.unary();
    for (;;) {
      const t = this.peek();
      if (!t || t.k !== 'op') break;
      if (t.v === '%') {
        this.i++;
        left = { t: 'pct', a: left };
        continue;
      }
      const prec = BIN_PREC[t.v];
      if (prec === undefined || prec < minPrec) break;
      this.i++;
      // «^» в Excel левоассоциативен, как и остальные операторы
      const right = this.expr(prec + 1);
      left = { t: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }

  private unary(): Node {
    const t = this.peek();
    if (t && t.k === 'op' && (t.v === '-' || t.v === '+')) {
      this.i++;
      const a = this.unary();
      return t.v === '-' ? { t: 'neg', a } : a;
    }
    return this.postfix(this.primary());
  }

  private postfix(node: Node): Node {
    let n = node;
    for (;;) {
      const t = this.peek();
      if (t && t.k === 'op' && t.v === '%') {
        this.i++;
        n = { t: 'pct', a: n };
      } else if (t && t.k === 'op' && t.v === '^') {
        // степень связывает сильнее унарного минуса слева: -2^2 = 4 в Excel,
        // поэтому обрабатываем её в expr; здесь ничего не делаем
        return n;
      } else return n;
    }
  }

  private primary(): Node {
    const t = this.toks[this.i++];
    if (!t) throw new FormulaSyntaxError('Формула оборвалась');
    switch (t.k) {
      case 'num':
        return { t: 'num', v: t.v };
      case 'str':
        return { t: 'str', v: t.v };
      case 'bool':
        return { t: 'bool', v: t.v };
      case 'err':
        return { t: 'err', v: t.v };
      case 'ref':
        return { t: 'ref', sheet: t.sheet, a: t.a, b: t.b };
      case 'name':
        return { t: 'name', name: t.name };
      case '(': {
        const inner = this.expr(0);
        const close = this.toks[this.i++];
        if (!close || close.k !== ')') throw new FormulaSyntaxError('Не хватает закрывающей скобки');
        return inner;
      }
      case 'fn': {
        const open = this.toks[this.i++];
        if (!open || open.k !== '(') throw new FormulaSyntaxError('Ожидается «(»');
        const args: Node[] = [];
        const next = this.peek();
        if (next && next.k === ')') {
          this.i++;
          return { t: 'fn', name: t.name, args };
        }
        for (;;) {
          const p = this.peek();
          // пропущенный аргумент: ЕСЛИ(A1;;2)
          if (p && (p.k === 'sep' || p.k === ')')) args.push({ t: 'str', v: '' });
          else args.push(this.expr(0));
          const sep = this.toks[this.i++];
          if (!sep) throw new FormulaSyntaxError('Не хватает закрывающей скобки');
          if (sep.k === ')') break;
          if (sep.k !== 'sep') throw new FormulaSyntaxError('Ожидается «;» между аргументами');
        }
        return { t: 'fn', name: t.name, args };
      }
      default:
        throw new FormulaSyntaxError('Неожиданный символ в формуле');
    }
  }
}

/**
 * Разбирает формулу. Сначала пробует нотацию с «;» если она есть, иначе с «,»,
 * а при ошибке — вторую нотацию (пользователь мог написать «=12,5*2»).
 */
export function parseFormula(src: string): Node {
  const semi = usesSemicolons(src);
  try {
    return new Parser(tokenize(src, semi)).parse();
  } catch (e) {
    if (semi) throw e;
    try {
      return new Parser(tokenize(src, true)).parse();
    } catch {
      throw e;
    }
  }
}

export { FormulaSyntaxError };

/** Токены формулы в той нотации, в которой она разбирается (как в parseFormula). */
export function formulaTokens(src: string): Tok[] {
  const semi = usesSemicolons(src);
  const first = tokenize(src, semi);
  try {
    new Parser(first).parse();
    return first;
  } catch (e) {
    if (semi) throw e;
    const second = tokenize(src, true);
    new Parser(second).parse();
    return second;
  }
}
