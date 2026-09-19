import { useState } from 'react';
import { parseLooseNumber } from '../formula/coerce';
import { numberToText } from '../formula/coerce';
import type { NamedValue } from '../model/types';
import { useUI } from '../ui/state';
import { gridApi } from './gridApi';
import { store, useStoreVersion } from './instance';

/** «КУРС_ЮАНЯ» → «Курс юаня» */
function pretty(name: string) {
  const s = name.replace(/_/g, ' ').toLocaleLowerCase('ru');
  return s.charAt(0).toLocaleUpperCase('ru') + s.slice(1);
}

function Param({ nv, index }: { nv: NamedValue; index: number }) {
  const [draft, setDraft] = useState<string | null>(null);
  const shown = typeof nv.value === 'number' ? numberToText(nv.value) : nv.value;
  const commit = () => {
    if (draft === null) return;
    const text = draft.trim();
    setDraft(null);
    if (text === shown) return;
    const num = parseLooseNumber(text);
    const names = [...store.meta.names];
    names[index] = { ...nv, value: num ?? text };
    store.setNames(names);
  };
  return (
    <label className="param" title={`${nv.name}${nv.note ? ' — ' + nv.note : ''}. В формулах: =…*${nv.name}`}>
      <span className="param-name">{pretty(nv.name)}</span>
      <input
        className="param-input"
        value={draft ?? shown}
        inputMode="decimal"
        spellCheck={false}
        size={Math.max(3, (draft ?? shown).length + 1)}
        onFocus={(e) => {
          setDraft(shown);
          requestAnimationFrame(() => e.target.select());
        }}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            (e.target as HTMLInputElement).blur();
            gridApi.focus();
          }
          if (e.key === 'Escape') {
            setDraft(null);
            gridApi.focus();
          }
        }}
      />
    </label>
  );
}

/** Именованные значения (курс, доставка) — видны и правятся над таблицей, как в первой строке Excel. */
export function ParamsBar() {
  useStoreVersion();
  const set = useUI((s) => s.set);
  const names = store.meta.names;
  return (
    <div className="params" role="group" aria-label="Параметры для формул">
      {names.slice(0, 4).map((nv, i) => (
        <Param key={nv.name} nv={nv} index={i} />
      ))}
      <button type="button" className="params-more" onClick={() => set({ dialog: { kind: 'settings' } })}>
        {names.length ? (names.length > 4 ? `ещё ${names.length - 4}` : 'изменить') : '+ параметр для формул'}
      </button>
    </div>
  );
}
