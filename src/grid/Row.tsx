import { memo, type CSSProperties } from 'react';
import { store } from '../app/instance';
import { safeHref } from '../model/format';
import type { CellDisplay } from '../model/store';
import type { Row, Sheet } from '../model/types';
import { useImageUrl } from '../ui/useImage';
import type { Geometry } from './geometry';

interface RowProps {
  geo: Geometry;
  sheet: Sheet;
  row: Row | undefined;
  /** видимая и физическая строка */
  vr: number;
  phys: number;
  y: number;
  h: number;
  /** диапазон видимых столбцов [c0, c1] */
  c0: number;
  c1: number;
  /** сдвиг x: для основной панели — frozenW, для левой — −RH_W */
  xShift: number;
  keyCol: number;
  /** эпоха вычислений — только для строк с формулами */
  calcKey: number;
  layoutKey: number;
}

function cellStyle(d: CellDisplay, left: number, width: number): CSSProperties {
  const st = d.style;
  const s: CSSProperties = { left, width };
  if (st.bg) s.backgroundColor = st.bg;
  if (st.fg) s.color = st.fg;
  if (st.b) s.fontWeight = 600;
  if (st.i) s.fontStyle = 'italic';
  if (st.u || st.s) s.textDecorationLine = [st.u ? 'underline' : '', st.s ? 'line-through' : ''].join(' ').trim();
  return s;
}

function hAlign(d: CellDisplay): string {
  if (d.style.ha) return d.style.ha;
  if (d.isError || typeof d.value === 'boolean') return 'center';
  return d.isNumber ? 'right' : 'left';
}

export const CellImage = memo(function CellImage({ id }: { id: string }) {
  const url = useImageUrl(id, 'thumb');
  if (url === null) return <span className="c-img c-img--missing" aria-label="Фото не найдено" />;
  if (!url) return <span className="c-img c-img--loading" aria-hidden />;
  return <img className="c-img" src={url} alt="" draggable={false} decoding="async" />;
});

/** Одна ячейка: обычная или объединённая (тогда она шире/выше и лежит поверх соседей). */
function renderCell(d: CellDisplay, c: number, vr: number, keyCol: number, left: number, width: number, merged = false) {
  const ha = hAlign(d);
  const va = d.style.va ?? (d.img ? 'middle' : 'middle');
  const wrap = !!d.style.wrap;
  const isKey = c === keyCol && d.text !== '';
  const href = d.href ? safeHref(d.href) : null;
  let content;
  if (isKey) {
    // вся ячейка артикула — кнопка карточки (ТЗ: «по тексту или самой ячейке»), здесь только подпись
    content = <span className="c-key">{d.text}</span>;
  } else if (href && d.text) {
    content = (
      <a className="c-link" href={href} target="_blank" rel="noopener noreferrer" tabIndex={-1} title={href}>
        {d.text}
      </a>
    );
  } else if (d.text) {
    content = <span className={d.isError ? 'c-t c-err' : 'c-t'}>{d.text}</span>;
  }
  return (
    <div
      key={c}
      className={
        'c' +
        (merged ? ' c--merged' : '') +
        (d.img ? ' c--img' : '') +
        (wrap ? ' c--wrap' : '') +
        (d.img && d.text ? ' c--both' : '') +
        (d.note ? ' c--note' : '') +
        (isKey ? ' c--key' : '')
      }
      title={d.note ?? (isKey ? 'Открыть карточку товара' : undefined)}
      data-open-card={isKey ? vr : undefined}
      data-ha={ha}
      data-va={va}
      style={cellStyle(d, left, width)}
    >
      {d.img && <CellImage id={d.img} />}
      {content}
    </div>
  );
}

export const GridRow = memo(function GridRow(p: RowProps) {
  const { geo, sheet, vr, phys, y, h, c0, c1, xShift, keyCol } = p;
  const cells = [];
  for (let vc = c0; vc <= c1; vc++) {
    const c = geo.cols[vc];
    const d = store.display(sheet, phys, c);
    cells.push(renderCell(d, c, vr, keyCol, geo.colX[vc] - xShift, geo.colX[vc + 1] - geo.colX[vc]));
  }
  return (
    <div className="gr" style={{ transform: `translateY(${y}px)`, height: h }} data-vr={vr}>
      {cells}
    </div>
  );
});

const BLANK = { text: '', img: undefined, note: undefined, href: undefined };

/**
 * Объединённые ячейки панели — поверх обычных строк. Непрозрачный фон закрывает ячейки под ними.
 * Если объединение пересекает границу закреплённых столбцов, его части рисуются в обеих панелях,
 * а значение — в той, где левая верхняя ячейка.
 */
export function MergeLayer({ geo, sheet, pane, r0, r1, c0, c1, keyCol }: { geo: Geometry; sheet: Sheet; pane: 'left' | 'main'; r0: number; r1: number; c0: number; c1: number; keyCol: number }) {
  const out = [];
  const xShift = pane === 'left' ? 0 : geo.frozenW;
  for (const b of geo.merges.boxes) {
    if (b.vr1 < r0 || b.vr0 > r1) continue;
    const a = pane === 'left' ? b.vc0 : Math.max(b.vc0, geo.frozen);
    const z = pane === 'left' ? Math.min(b.vc1, geo.frozen - 1) : b.vc1;
    if (a > z || (pane === 'main' && (z < c0 || a > c1))) continue;
    const y = geo.rowY[b.vr0];
    const d = store.display(sheet, b.phys, b.col);
    out.push(
      <div key={b.merge.r0 + b.merge.c0} className="gr gr--merge" style={{ transform: `translateY(${y}px)`, height: geo.rowY[b.vr1 + 1] - y }}>
        {renderCell(a === b.vc0 ? d : { ...d, ...BLANK }, b.col, b.vr0, keyCol, geo.colX[a] - xShift, geo.colX[z + 1] - geo.colX[a], true)}
      </div>,
    );
  }
  return <>{out}</>;
}
