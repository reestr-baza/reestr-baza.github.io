import { viewMerges, type ViewMerges } from '../model/merges';
import type { SheetView, Store } from '../model/store';
import type { Sheet } from '../model/types';

export const RH_W = 56;
export const HD_H = 34;

export interface Geometry {
  sheet: Sheet;
  view: SheetView;
  /** видимый столбец → физический индекс */
  cols: number[];
  /** левые края видимых столбцов; colX[n] — общая ширина */
  colX: Float64Array;
  frozen: number;
  frozenW: number;
  /** верхние края видимых строк; rowY[n] — общая высота */
  rowY: Float64Array;
  totalW: number;
  totalH: number;
  /** видимая строка → физический индекс */
  phys: Int32Array;
  /** физический индекс → видимая строка (−1, если скрыта) */
  vrOf: Int32Array;
  /** физический столбец → видимый (−1, если скрыт) */
  vcOf: Int32Array;
  /** видимые строки, перед которыми есть скрытые вручную */
  hiddenBefore: Set<number>;
  /** видимые столбцы, перед которыми есть скрытые */
  hiddenColBefore: Set<number>;
  /** есть скрытые строки/столбцы в конце */
  hiddenAfterRows: boolean;
  hiddenAfterCols: boolean;
  /** Объединённые ячейки в координатах представления */
  merges: ViewMerges;
}

/** maxFrozenW — сколько места могут занять закреплённые столбцы (на телефоне меньше). */
export function buildGeometry(store: Store, sheet: Sheet, view: SheetView, maxFrozenW = Infinity): Geometry {
  const cols = view.cols;
  const colX = new Float64Array(cols.length + 1);
  const hiddenColBefore = new Set<number>();
  let prevPhys = -1;
  for (let i = 0; i < cols.length; i++) {
    colX[i + 1] = colX[i] + sheet.columns[cols[i]].w;
    if (cols[i] - prevPhys > 1) hiddenColBefore.add(i);
    prevPhys = cols[i];
  }
  const hiddenAfterCols = prevPhys < sheet.columns.length - 1;
  let frozen = Math.min(sheet.frozen, cols.length - 1 < 0 ? 0 : cols.length - 1);
  while (frozen > 0 && colX[frozen] > maxFrozenW) frozen--;
  const frozenW = colX[frozen];

  const n = view.rows.length;
  const rowY = new Float64Array(n + 1);
  const phys = new Int32Array(n);
  const hiddenBefore = new Set<number>();
  let prev = -1;
  for (let i = 0; i < n; i++) {
    const id = view.rows[i];
    const row = sheet.rows.get(id);
    rowY[i + 1] = rowY[i] + store.rowHeight(sheet, row);
    const p = view.rowIndex.get(id)!;
    phys[i] = p;
    if (p - prev > 1) {
      // между соседними видимыми строками есть скрытые — отметим, если хоть одна скрыта вручную
      for (let k = prev + 1; k < p; k++) {
        if (sheet.rows.get(sheet.rowOrder[k])?.hidden) {
          hiddenBefore.add(i);
          break;
        }
      }
    }
    prev = p;
  }
  const vrOf = new Int32Array(sheet.rowOrder.length).fill(-1);
  for (let i = 0; i < n; i++) vrOf[phys[i]] = i;
  const vcOf = new Int32Array(sheet.columns.length).fill(-1);
  cols.forEach((c, i) => (vcOf[c] = i));
  let hiddenAfterRows = false;
  for (let k = prev + 1; k < sheet.rowOrder.length; k++) {
    if (sheet.rows.get(sheet.rowOrder[k])?.hidden) {
      hiddenAfterRows = true;
      break;
    }
  }
  return {
    sheet,
    view,
    cols,
    colX,
    frozen,
    frozenW,
    rowY,
    totalW: colX[cols.length],
    totalH: rowY[n],
    phys,
    vrOf,
    vcOf,
    hiddenBefore,
    hiddenColBefore,
    hiddenAfterRows,
    hiddenAfterCols,
    merges: viewMerges(store, sheet, view),
  };
}

/** Индекс i с prefix[i] ≤ pos < prefix[i+1], ограниченный [0, n−1]. */
export function indexAt(prefix: Float64Array, pos: number): number {
  const n = prefix.length - 1;
  if (n <= 0) return 0;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (prefix[mid] <= pos) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}
