import { X } from 'lucide-react';
import { useEffect, useId, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

interface Props {
  title: ReactNode;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  className?: string;
  /** Скрыть заголовок визуально (есть свой) */
  bare?: boolean;
  initialFocus?: string;
  labelText?: string;
}

const FOCUSABLE = 'a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])';

export function Dialog({ title, onClose, children, footer, className, bare, initialFocus, labelText }: Props) {
  const ref = useRef<HTMLDivElement>(null);
  const titleId = useId();
  const closeRef = useRef(onClose);
  closeRef.current = onClose;

  useEffect(() => {
    const prev = document.activeElement as HTMLElement | null;
    const root = document.getElementById('root');
    root?.setAttribute('inert', '');
    const el = ref.current!;
    const first = (initialFocus && el.querySelector<HTMLElement>(initialFocus)) || el.querySelector<HTMLElement>('[data-autofocus]') || el.querySelector<HTMLElement>(FOCUSABLE);
    first?.focus({ preventScroll: true });
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !e.defaultPrevented) {
        // Escape в открытом поповере закроет только его
        if (document.querySelector('.pop')) return;
        e.preventDefault();
        closeRef.current();
      }
      if (e.key === 'Tab') {
        const items = Array.from(el.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((x) => x.offsetParent !== null);
        if (!items.length) return;
        const i = items.indexOf(document.activeElement as HTMLElement);
        if (e.shiftKey && (i === 0 || i === -1)) {
          e.preventDefault();
          items[items.length - 1].focus();
        } else if (!e.shiftKey && i === items.length - 1) {
          e.preventDefault();
          items[0].focus();
        }
      }
    };
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('keydown', onKey);
      // вложенные диалоги: снимаем inert, только если это последний
      if (!document.querySelector('.dlg-backdrop ~ .dlg-backdrop')) root?.removeAttribute('inert');
      if (prev && document.contains(prev)) prev.focus({ preventScroll: true });
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return createPortal(
    <div
      className="dlg-backdrop"
      onPointerDown={(e) => {
        if (e.target === e.currentTarget) closeRef.current();
      }}
    >
      <div ref={ref} className={'dlg ' + (className ?? '')} role="dialog" aria-modal="true" aria-labelledby={labelText ? undefined : titleId} aria-label={labelText}>
        {!bare && (
          <header className="dlg-head">
            <h2 id={titleId} className="dlg-title">
              {title}
            </h2>
            <button type="button" className="icon-btn" onClick={onClose} aria-label="Закрыть">
              <X size={18} strokeWidth={1.75} />
            </button>
          </header>
        )}
        <div className="dlg-body">{children}</div>
        {footer && <footer className="dlg-foot">{footer}</footer>}
      </div>
    </div>,
    document.body,
  );
}
