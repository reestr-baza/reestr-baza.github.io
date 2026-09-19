import { computePosition, flip, offset, shift } from '@floating-ui/dom';
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

/**
 * Одна общая подсказка для всех [data-tip]: первая появляется с задержкой,
 * следующие при переходе по панели — сразу, без анимации.
 */
export function Tooltips() {
  const [tip, setTip] = useState<{ text: string; el: HTMLElement } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  const timer = useRef(0);
  const lastHide = useRef(0);

  useEffect(() => {
    if (!matchMedia('(hover: hover) and (pointer: fine)').matches) return;
    const show = (el: HTMLElement) => {
      const text = el.dataset.tip;
      if (!text) return;
      window.clearTimeout(timer.current);
      const warm = Date.now() - lastHide.current < 500;
      timer.current = window.setTimeout(() => setTip({ text, el }), warm ? 0 : 500);
    };
    const hide = () => {
      window.clearTimeout(timer.current);
      setTip((t) => {
        if (t) lastHide.current = Date.now();
        return null;
      });
    };
    const over = (e: PointerEvent) => {
      if (e.pointerType !== 'mouse') return;
      const el = (e.target as HTMLElement).closest<HTMLElement>('[data-tip]');
      if (el) show(el);
      else hide();
    };
    document.addEventListener('pointerover', over);
    document.addEventListener('pointerdown', hide, true);
    document.addEventListener('keydown', hide, true);
    window.addEventListener('blur', hide);
    return () => {
      document.removeEventListener('pointerover', over);
      document.removeEventListener('pointerdown', hide, true);
      document.removeEventListener('keydown', hide, true);
      window.removeEventListener('blur', hide);
    };
  }, []);

  useEffect(() => {
    if (!tip || !ref.current) return;
    const el = ref.current;
    void computePosition(tip.el, el, { placement: 'bottom', strategy: 'fixed', middleware: [offset(6), flip(), shift({ padding: 8 })] }).then(({ x, y }) => {
      el.style.left = `${x}px`;
      el.style.top = `${y}px`;
    });
  }, [tip]);

  if (!tip) return null;
  return createPortal(
    <div ref={ref} className="tip" role="tooltip" style={{ left: -9999, top: -9999 }}>
      {tip.text}
    </div>,
    document.body,
  );
}
