import { useEffect, useRef, type ReactNode } from 'react';
import { Popover, type Anchor } from './Popover';

export type MenuItem =
  | {
      label: string;
      icon?: ReactNode;
      hint?: string;
      onSelect: () => void;
      disabled?: boolean;
      danger?: boolean;
      checked?: boolean;
    }
  | 'sep'
  | { heading: string };

export function MenuList({ items, onClose, autoFocus = true }: { items: MenuItem[]; onClose: () => void; autoFocus?: boolean }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (autoFocus) ref.current?.querySelector<HTMLButtonElement>('[role^="menuitem"]:not([disabled])')?.focus({ preventScroll: true });
  }, [autoFocus]);

  const onKeyDown = (e: React.KeyboardEvent) => {
    const btns = Array.from(ref.current!.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]:not([disabled])'));
    const i = btns.indexOf(document.activeElement as HTMLButtonElement);
    let next = -1;
    if (e.key === 'ArrowDown') next = (i + 1) % btns.length;
    else if (e.key === 'ArrowUp') next = (i - 1 + btns.length) % btns.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = btns.length - 1;
    else if (e.key === 'Tab') {
      e.preventDefault();
      onClose();
      return;
    }
    if (next >= 0) {
      e.preventDefault();
      btns[next]?.focus();
    }
  };

  return (
    <div ref={ref} className="menu" role="menu" onKeyDown={onKeyDown}>
      {items.map((it, i) => {
        if (it === 'sep') return <div key={i} className="menu-sep" role="separator" />;
        if ('heading' in it)
          return (
            <div key={i} className="menu-head" role="presentation">
              {it.heading}
            </div>
          );
        return (
          <button
            key={i}
            type="button"
            role={it.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'}
            aria-checked={it.checked}
            className={'menu-item' + (it.danger ? ' is-danger' : '')}
            disabled={it.disabled}
            onClick={() => {
              onClose();
              it.onSelect();
            }}
          >
            <span className="menu-ic" aria-hidden>
              {it.checked ? '✓' : it.icon}
            </span>
            <span className="menu-label">{it.label}</span>
            {it.hint && <span className="menu-hint">{it.hint}</span>}
          </button>
        );
      })}
    </div>
  );
}

export function Menu({ anchor, items, onClose, label, placement }: { anchor: Anchor; items: MenuItem[]; onClose: () => void; label?: string; placement?: 'bottom-start' | 'bottom-end' | 'right-start' }) {
  return (
    <Popover anchor={anchor} onClose={onClose} className="pop--menu" label={label} placement={placement ?? 'bottom-start'} offsetPx={4}>
      <MenuList items={items} onClose={onClose} />
    </Popover>
  );
}
