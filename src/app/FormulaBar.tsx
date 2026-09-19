import { useEffect, useRef, useState } from 'react';
import { colToLetters, lettersToCol, normalizeLookalikes } from '../formula/a1';
import { shiftFormula } from '../formula/refs';
import { applyHint, FunctionHints, useFunctionHints } from '../grid/Editor';
import { useUI } from '../ui/state';
import { applyColumnFormula, cancelEdit, cellAt, commitEdit, ctx, editableText, setSelection, startEdit } from './actions';
import { gridApi } from './gridApi';
import { useStoreVersion } from './instance';

function NameBox() {
  const sel = useUI((s) => s.sel);
  const x = cellAt(sel.ar, sel.ac);
  const multi = sel.ar !== sel.fr || sel.ac !== sel.fc;
  const label = x ? `${colToLetters(x.c)}${x.phys + 1}` : '';
  const [draft, setDraft] = useState<string | null>(null);
  const range = multi ? `${Math.abs(sel.fr - sel.ar) + 1}×${Math.abs(sel.fc - sel.ac) + 1}` : '';
  return (
    <input
      className="namebox"
      aria-label="Адрес ячейки — введите адрес и нажмите Enter, чтобы перейти"
      value={draft ?? (multi ? range : label)}
      spellCheck={false}
      onFocus={(e) => {
        setDraft(label);
        requestAnimationFrame(() => e.target.select());
      }}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => setDraft(null)}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          setDraft(null);
          gridApi.focus();
        }
        if (e.key !== 'Enter') return;
        const m = /^([A-Za-zА-Яа-я]{1,3})(\d+)$/.exec(normalizeLookalikes((draft ?? '').trim()));
        if (m) {
          const { sheet, view } = ctx();
          const c = lettersToCol(m[1]);
          const phys = parseInt(m[2], 10) - 1;
          const vc = view.cols.indexOf(c);
          const vr = view.rows.indexOf(sheet.rowOrder[phys]);
          if (vc >= 0 && vr >= 0) setSelection({ ar: vr, ac: vc, fr: vr, fc: vc });
          else useUI.getState().toast({ text: 'Ячейка скрыта фильтром или не существует', tone: 'error' });
        }
        setDraft(null);
        gridApi.focus();
      }}
    />
  );
}

export function FormulaBar() {
  useStoreVersion();
  const sel = useUI((s) => s.sel);
  const edit = useUI((s) => s.edit);
  const set = useUI((s) => s.set);
  const ref = useRef<HTMLTextAreaElement>(null);
  const [caret, setCaret] = useState(0);
  const { sheet } = ctx();
  const x = cellAt(sel.ar, sel.ac);
  const cell = x?.row?.cells[x.colId];
  const fromColumn = !!(x && x.col.formula && !cell?.f && !(cell && (cell.v !== undefined || cell.img)));
  const text = edit ? edit.text : editableText(sel.ar, sel.ac);
  const hints = useFunctionHints(edit?.source === 'bar' ? edit.text : '', caret);

  useEffect(() => {
    const ta = ref.current;
    if (ta && edit?.source === 'bar' && edit.caret !== undefined && document.activeElement === ta) {
      ta.setSelectionRange(edit.caret, edit.caret);
      setCaret(edit.caret);
    }
  }, [edit?.caret, edit?.source, edit?.text]);

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (hints.list.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const d = e.key === 'ArrowDown' ? 1 : -1;
        hints.setActive((hints.active + d + hints.list.length) % hints.list.length);
        return;
      }
      if (e.key === 'Tab' && edit) {
        e.preventDefault();
        const r = applyHint(edit.text, e.currentTarget.selectionStart, hints.prefix!, hints.list[hints.active]);
        set({ edit: { ...edit, text: r.text, caret: r.caret } });
        return;
      }
    }
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      commitEdit(e.shiftKey ? 'up' : 'down');
    } else if (e.key === 'Enter' && e.altKey && edit) {
      e.preventDefault();
      const s = e.currentTarget.selectionStart;
      set({ edit: { ...edit, text: edit.text.slice(0, s) + '\n' + edit.text.slice(e.currentTarget.selectionEnd), caret: s + 1 } });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commitEdit(e.shiftKey ? 'left' : 'right');
    }
  };

  const lines = Math.min(4, text.split('\n').length);

  return (
    <div className="fbar">
      <NameBox />
      <span className="fbar-fx" aria-hidden>
        ƒ
      </span>
      <div className="fbar-field">
        <textarea
          ref={ref}
          className={'fbar-input' + (text.startsWith('=') ? ' is-formula' : '')}
          rows={lines}
          aria-label="Содержимое ячейки"
          value={text}
          spellCheck={false}
          onFocus={() => {
            if (!edit) startEdit('edit', undefined, 'bar');
            else if (edit.source !== 'bar') set({ edit: { ...edit, source: 'bar' } });
          }}
          onChange={(e) => {
            const cur = useUI.getState().edit;
            setCaret(e.target.selectionStart);
            if (cur) set({ edit: { ...cur, text: e.target.value, caret: undefined, source: 'bar' } });
            else startEdit('edit', e.target.value, 'bar');
          }}
          onSelect={(e) => setCaret(e.currentTarget.selectionStart)}
          onKeyDown={onKeyDown}
        />
        {edit?.source === 'bar' && (
          <FunctionHints
            list={hints.list}
            active={hints.active}
            style={{ left: 0, top: '100%' }}
            onPick={(i) => {
              const r = applyHint(edit.text, caret, hints.prefix!, hints.list[i]);
              set({ edit: { ...edit, text: r.text, caret: r.caret } });
              ref.current?.focus();
            }}
          />
        )}
      </div>
      {fromColumn && !edit && x && (
        <div className="fbar-colf">
          <span className="fbar-colf-tag" title={`Для каждой строки: ${shiftFormula(x.col.formula!, x.phys, 0)}`}>
            Формула столбца
          </span>
          <button type="button" className="btn btn--sm btn--ghost" onClick={() => applyColumnFormula(x.colId, undefined)}>
            Убрать
          </button>
        </div>
      )}
      {!fromColumn && !edit && x && cell?.f && !x.col.formula && sheet.rowOrder.length > 2 && (
        <button
          type="button"
          className="btn btn--sm btn--ghost fbar-apply"
          onClick={() => applyColumnFormula(x.colId, shiftFormula(cell.f!, -x.phys, 0))}
          title="Одна формула на весь столбец — новые строки посчитаются сами"
        >
          Применить ко всему столбцу
        </button>
      )}
    </div>
  );
}

