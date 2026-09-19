import type { EvalCtx } from './evaluate';
import type { RangeVal } from './values';

export function makeRange(ctx: EvalCtx, sheetId: string, r1: number, c1: number, r2: number, c2: number): RangeVal {
  const top = Math.min(r1, r2);
  const left = Math.min(c1, c2);
  const host = ctx.host;
  return {
    kind: 'range',
    rows: Math.abs(r2 - r1) + 1,
    cols: Math.abs(c2 - c1) + 1,
    top,
    left,
    sheetId,
    get: (r, c) => host.cellValue(sheetId, top + r, left + c),
  };
}
