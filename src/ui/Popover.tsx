import { autoUpdate, computePosition, flip, offset, shift, size, type Placement } from '@floating-ui/dom';
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

export type Anchor = { x: number; y: number } | { rect: { left: number; top: number; right: number; bottom: number } } | { el: HTMLElement };

function virtualEl(anchor: Anchor) {
  if ('el' in anchor) return anchor.el;
  const r = 'rect' in anchor ? anchor.rect : { left: anchor.x, top: anchor.y, right: anchor.x, bottom: anchor.y };
  return {
    getBoundingClientRect: () =>
      ({ x: r.left, y: r.top, left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.right - r.left, height: r.bottom - r.top }) as DOMRect,
  };
}

interface Props {
  anchor: Anchor;
  placement?: Placement;
  onClose: () => void;
  children: ReactNode;
  className?: string;
  role?: string;
  label?: string;
  /** Не закрывать по клику вне (например, пока открыт вложенный выбор цвета) */
  modalLike?: boolean;
  offsetPx?: number;
}

export function Popover({ anchor, placement = 'bottom-start', onClose, children, className, role, label, offsetPx = 6 }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  // запоминаем фокус ещё при рендере: вложенное меню переведёт его на свой пункт раньше наших эффектов
  const [opener] = useState(() => document.activeElement as HTMLElement | null);

  useLayoutEffect(() => {
    const el = ref.current!;
    const ref0 = virtualEl(anchor);
    const update = () =>
      computePosition(ref0 as Element, el, {
        placement,
        strategy: 'fixed',
        middleware: [
          offset(offsetPx),
          flip({ padding: 8 }),
          shift({ padding: 8 }),
          size({
            padding: 8,
            apply({ availableHeight }) {
              el.style.maxHeight = `${Math.max(160, availableHeight)}px`;
            },
          }),
        ],
      }).then(({ x, y, placement: p }) => {
        el.style.left = `${x}px`;
        el.style.top = `${y}px`;
        el.dataset.side = p.split('-')[0];
      });
    if ('el' in anchor) return autoUpdate(anchor.el, el, update);
    void update();
    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, ['el' in anchor ? anchor.el : JSON.stringify(anchor), placement]);

  useEffect(() => {
    const prev = opener;
    const onDown = (e: PointerEvent) => {
      const t = e.target as Node;
      if (ref.current?.contains(t)) return;
      // клики по вложенным поповерам (выбор цвета внутри меню) не закрывают родителя
      if ((t as HTMLElement).closest?.('[data-popover]')) return;
      closeRef.current();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        closeRef.current();
      }
    };
    const t = window.setTimeout(() => document.addEventListener('pointerdown', onDown, true), 0);
    document.addEventListener('keydown', onKey, true);
    return () => {
      window.clearTimeout(t);
      document.removeEventListener('pointerdown', onDown, true);
      document.removeEventListener('keydown', onKey, true);
      // пункт меню уже удалён из DOM — фокус «выпал» на body; возвращаем его туда, откуда открыли (обычно таблица)
      const ae = document.activeElement;
      const lost = !ae || ae === document.body || !!ref.current?.contains(ae);
      if (prev && document.contains(prev) && lost) prev.focus({ preventScroll: true });
    };
  }, []);

  return createPortal(
    <div ref={ref} className={'pop ' + (className ?? '')} role={role} aria-label={label} data-popover style={{ position: 'fixed', left: -9999, top: -9999 }}>
      {children}
    </div>,
    document.body,
  );
}
