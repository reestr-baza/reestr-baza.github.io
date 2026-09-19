import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { cancelEdit, commitEdit } from '../app/actions';
import { FUNCTION_HINTS } from '../formula/functions';
import { useUI } from '../ui/state';
import type { Geometry } from './geometry';
import { isFormulaPointMode } from './keyboard';

let measureCtx: CanvasRenderingContext2D | null = null;
function textWidth(s: string): number {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  if (!measureCtx) return s.length * 7;
  measureCtx.font = '13px "Golos Text Variable", system-ui, sans-serif';
  return measureCtx.measureText(s).width;
}

/** Имя функции, которое пользователь сейчас набирает перед курсором. */
export function typedFunctionPrefix(text: string, caret: number): string | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const m = /(?:^=|[=(;,+\-*/^&<>\s])([A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9.]*)$/.exec(before);
  if (!m) return null;
  // внутри строки подсказки не нужны
  const quotes = (before.match(/"/g) ?? []).length;
  if (quotes % 2) return null;
  return m[1];
}

export function functionSuggestions(prefix: string | null): [string, string, string][] {
  if (!prefix || prefix.length < 1) return [];
  const up = prefix.toUpperCase();
  return FUNCTION_HINTS.filter(([ru, en]) => ru.startsWith(up) || en.startsWith(up)).slice(0, 7);
}

/** Подсказки функций под полем ввода. Общие для ячейки и строки формул. */
export function useFunctionHints(text: string, caret: number) {
  const prefix = typedFunctionPrefix(text, caret);
  const list = useMemo(() => functionSuggestions(prefix), [prefix]);
  const [active, setActive] = useState(0);
  useEffect(() => setActive(0), [prefix]);
  return { prefix, list, active, setActive };
}

export function applyHint(text: string, caret: number, prefix: string, hint: [string, string, string]) {
  const useRu = /[А-Яа-яЁё]/.test(prefix) || !/^[A-Za-z]/.test(prefix);
  const name = useRu ? hint[0] : hint[1];
  const start = caret - prefix.length;
  const next = text.slice(0, start) + name + '(' + text.slice(caret);
  return { text: next, caret: start + name.length + 1 };
}

export function FunctionHints({
  list,
  active,
  onPick,
  style,
}: {
  list: [string, string, string][];
  active: number;
  onPick: (i: number) => void;
  style?: React.CSSProperties;
}) {
  if (!list.length) return null;
  return (
    <ul className="fn-hints" role="listbox" aria-label="Функции" style={style}>
      {list.map((h, i) => (
        <li
          key={h[0]}
          role="option"
          aria-selected={i === active}
          className={i === active ? 'is-active' : ''}
          onPointerDown={(e) => {
            e.preventDefault();
            onPick(i);
          }}
        >
          <span className="fn-name">{h[0]}</span>
          {h[1] !== h[0] && <span className="fn-en">{h[1]}</span>}
          <span className="fn-desc">{h[2]}</span>
        </li>
      ))}
      <li className="fn-foot" aria-hidden>
        Tab — вставить
      </li>
    </ul>
  );
}

export function CellEditor({ geo, xShift }: { geo: Geometry; xShift: number; pane: 'left' | 'main' }) {
  const edit = useUI((s) => s.edit)!;
  const set = useUI((s) => s.set);
  const ref = useRef<HTMLTextAreaElement>(null);
  const [caret, setCaret] = useState(edit.text.length);
  const hints = useFunctionHints(edit.text, caret);

  useLayoutEffect(() => {
    const ta = ref.current;
    if (!ta || edit.source !== 'cell') return;
    ta.focus({ preventScroll: true });
    const end = ta.value.length;
    ta.setSelectionRange(end, end);
    setCaret(end);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [edit.r, edit.c]);

  // курсор после вставки ссылки мышью
  useLayoutEffect(() => {
    const ta = ref.current;
    if (ta && edit.caret !== undefined && document.activeElement === ta) {
      ta.setSelectionRange(edit.caret, edit.caret);
      setCaret(edit.caret);
    }
  }, [edit.caret, edit.text]);

  const x = geo.colX[edit.c] - xShift;
  const y = geo.rowY[edit.r];
  const cellW = geo.colX[edit.c + 1] - geo.colX[edit.c];
  const cellH = geo.rowY[edit.r + 1] - geo.rowY[edit.r];
  const lines = edit.text.split('\n');
  const widest = Math.max(...lines.map(textWidth));
  const w = Math.max(cellW + 1, Math.min(520, widest + 28));
  const h = Math.max(cellH + 1, lines.length * 19 + 12);

  const update = (text: string, c: number) => {
    set({ edit: { ...edit, text, caret: undefined } });
    setCaret(c);
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ta = e.currentTarget;
    if (hints.list.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const d = e.key === 'ArrowDown' ? 1 : -1;
        hints.setActive((hints.active + d + hints.list.length) % hints.list.length);
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        const r = applyHint(edit.text, ta.selectionStart, hints.prefix!, hints.list[hints.active]);
        set({ edit: { ...edit, text: r.text, caret: r.caret } });
        setCaret(r.caret);
        return;
      }
    }
    if (e.key === 'Enter') {
      if (e.altKey) {
        e.preventDefault();
        const s = ta.selectionStart;
        const text = edit.text.slice(0, s) + '\n' + edit.text.slice(ta.selectionEnd);
        set({ edit: { ...edit, text, caret: s + 1 } });
        return;
      }
      e.preventDefault();
      commitEdit(e.shiftKey ? 'up' : 'down');
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit(e.shiftKey ? 'left' : 'right');
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit();
      return;
    }
    if (e.key === 'F2') {
      e.preventDefault();
      set({ edit: { ...edit, mode: edit.mode === 'enter' ? 'edit' : 'enter' } });
      return;
    }
    if (edit.mode === 'enter' && e.key.startsWith('Arrow') && !e.shiftKey) {
      if (isFormulaPointMode(edit.text, ta.selectionStart)) return;
      e.preventDefault();
      const dir = { ArrowDown: 'down', ArrowUp: 'up', ArrowLeft: 'left', ArrowRight: 'right' } as const;
      commitEdit(dir[e.key as keyof typeof dir]);
    }
  };

  const isFormula = edit.text.startsWith('=');
  return (
    <>
      <textarea
        ref={ref}
        className={'ed' + (isFormula ? ' ed--formula' : '')}
        style={{ left: x - 1, top: y - 1, width: w, height: h }}
        value={edit.text}
        spellCheck={false}
        aria-label="Значение ячейки"
        onChange={(e) => update(e.target.value, e.target.selectionStart)}
        onSelect={(e) => setCaret(e.currentTarget.selectionStart)}
        onKeyDown={onKeyDown}
        onFocus={() => edit.source !== 'cell' && set({ edit: { ...edit, source: 'cell' } })}
      />
      {edit.source === 'cell' && (
        <FunctionHints
          list={hints.list}
          active={hints.active}
          style={{ left: x - 1, top: y + h + 2 }}
          onPick={(i) => {
            const r = applyHint(edit.text, caret, hints.prefix!, hints.list[i]);
            set({ edit: { ...edit, text: r.text, caret: r.caret } });
            setCaret(r.caret);
            ref.current?.focus();
          }}
        />
      )}
    </>
  );
}
