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

export const GridRow = memo(function GridRow(p: RowProps) {
  const { geo, sheet, vr, phys, y, h, c0, c1, xShift, keyCol } = p;
  const cells = [];
  for (let vc = c0; vc <= c1; vc++) {
    const c = geo.cols[vc];
    const d = store.display(sheet, phys, c);
    const left = geo.colX[vc] - xShift;
    const width = geo.colX[vc + 1] - geo.colX[vc];
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
    cells.push(
      <div
        key={c}
        className={
          'c' + (d.img ? ' c--img' : '') + (wrap ? ' c--wrap' : '') + (d.img && d.text ? ' c--both' : '') + (d.note ? ' c--note' : '') + (isKey ? ' c--key' : '')
        }
        title={d.note ?? (isKey ? 'Открыть карточку товара' : undefined)}
        data-open-card={isKey ? vr : undefined}
        data-ha={ha}
        data-va={va}
        style={cellStyle(d, left, width)}
      >
        {d.img && <CellImage id={d.img} />}
        {content}
      </div>,
    );
  }
  return (
    <div className="gr" style={{ transform: `translateY(${y}px)`, height: h }} data-vr={vr}>
      {cells}
    </div>
  );
});
