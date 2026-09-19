import type { SheetView, Store } from './store';
import type { Merge, RowId, Sheet } from './types';

/** Объединение в координатах текущего представления (видимые строки и столбцы). */
export interface MergeBox {
  vr0: number;
  vr1: number;
  vc0: number;
  vc1: number;
  /** Левая верхняя ячейка: физическая строка и столбец — её значение и оформление показывает объединение */
  phys: number;
  col: number;
  merge: Merge;
}

export interface ViewMerges {
  boxes: MergeBox[];
  /** Объединение, которое накрывает видимую ячейку */
  at(vr: number, vc: number): MergeBox | undefined;
}

const NONE: ViewMerges = { boxes: [], at: () => undefined };
const cache = new WeakMap<SheetView, { merges: Merge[] | undefined; columns: unknown; res: ViewMerges }>();

/** Физические границы объединения; null — угол удалён. */
export function mergeBounds(store: Store, sheet: Sheet, m: Merge, rowIndex?: Map<RowId, number>) {
  const ri = (id: RowId) => (rowIndex ? rowIndex.get(id) : sheet.rowOrder.indexOf(id)) ?? -1;
  const a = ri(m.r0);
  const b = ri(m.r1);
  const c = store.colIndexOf(sheet, m.c0);
  const d = store.colIndexOf(sheet, m.c1);
  if (a < 0 || b < 0 || c < 0 || d < 0) return null;
  return { p0: Math.min(a, b), p1: Math.max(a, b), q0: Math.min(c, d), q1: Math.max(c, d) };
}

/**
 * Объединения листа в текущем представлении. Скрытые и отфильтрованные строки внутри объединения
 * просто не занимают места — объединение сжимается до видимых строк (фильтр не меняет порядок строк).
 */
export function viewMerges(store: Store, sheet: Sheet, view: SheetView): ViewMerges {
  if (!sheet.merges?.length) return NONE;
  const hit = cache.get(view);
  if (hit && hit.merges === sheet.merges && hit.columns === sheet.columns) return hit.res;

  const vrOf = new Int32Array(sheet.rowOrder.length).fill(-1);
  view.rows.forEach((id, i) => {
    const p = view.rowIndex.get(id);
    if (p !== undefined) vrOf[p] = i;
  });
  const vcOf = new Int32Array(sheet.columns.length).fill(-1);
  view.cols.forEach((c, i) => (vcOf[c] = i));

  const boxes: MergeBox[] = [];
  const index = new Map<number, MergeBox>();
  for (const m of sheet.merges) {
    const b = mergeBounds(store, sheet, m, view.rowIndex);
    if (!b) continue;
    let vr0 = -1;
    let vr1 = -1;
    for (let p = b.p0; p <= b.p1; p++) {
      if (vrOf[p] < 0) continue;
      if (vr0 < 0) vr0 = vrOf[p];
      vr1 = vrOf[p];
    }
    let vc0 = -1;
    let vc1 = -1;
    for (let q = b.q0; q <= b.q1; q++) {
      if (vcOf[q] < 0) continue;
      if (vc0 < 0) vc0 = vcOf[q];
      vc1 = vcOf[q];
    }
    if (vr0 < 0 || vc0 < 0) continue;
    // в представлении видимые строки идут подряд, только если порядок не нарушен
    if (vr1 - vr0 > b.p1 - b.p0) continue;
    const box: MergeBox = { vr0, vr1, vc0, vc1, phys: b.p0, col: b.q0, merge: m };
    boxes.push(box);
    for (let r = vr0; r <= vr1; r++) for (let c = vc0; c <= vc1; c++) index.set(r * 1024 + c, box);
  }
  const res: ViewMerges = boxes.length ? { boxes, at: (vr, vc) => index.get(vr * 1024 + vc) } : NONE;
  cache.set(view, { merges: sheet.merges, columns: sheet.columns, res });
  return res;
}

/** Расширить прямоугольник так, чтобы объединения попадали в него целиком (как выделение в Excel). */
export function growOverMerges(vm: ViewMerges, r: { r1: number; r2: number; c1: number; c2: number }) {
  let { r1, r2, c1, c2 } = r;
  let changed = vm.boxes.length > 0;
  while (changed) {
    changed = false;
    for (const b of vm.boxes) {
      if (b.vr1 < r1 || b.vr0 > r2 || b.vc1 < c1 || b.vc0 > c2) continue;
      if (b.vr0 < r1) (r1 = b.vr0), (changed = true);
      if (b.vr1 > r2) (r2 = b.vr1), (changed = true);
      if (b.vc0 < c1) (c1 = b.vc0), (changed = true);
      if (b.vc1 > c2) (c2 = b.vc1), (changed = true);
    }
  }
  return { r1, r2, c1, c2 };
}
