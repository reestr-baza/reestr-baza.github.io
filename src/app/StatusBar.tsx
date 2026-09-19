import { Plus } from 'lucide-react';
import { useMemo, useState } from 'react';
import { formatNumber } from '../model/format';
import { emptySheet } from '../model/store';
import { selRect, useUI } from '../ui/state';
import { cellAt, ctx } from './actions';
import { gridApi } from './gridApi';
import { store, useSaveState, useStoreVersion } from './instance';

function plural(n: number, one: string, few: string, many: string) {
  const m10 = n % 10;
  const m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
  return many;
}

export function rowsWord(n: number) {
  return plural(n, 'строка', 'строки', 'строк');
}

function SheetTabs() {
  useStoreVersion();
  const set = useUI((s) => s.set);
  const [renaming, setRenaming] = useState<string | null>(null);
  const [draft, setDraft] = useState('');
  const sheets = store.sheetList();
  const active = store.meta.activeSheet;

  const switchTo = (id: string) => {
    set({ edit: null, filterMenu: null, sel: { ar: 0, ac: 0, fr: 0, fc: 0 } });
    store.setActiveSheet(id);
    gridApi.focus();
  };

  return (
    <div className="tabs" role="tablist" aria-label="Листы">
      {sheets.map((s) =>
        renaming === s.id ? (
          <input
            key={s.id}
            className="tab tab-input"
            autoFocus
            aria-label="Название листа"
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={() => {
              if (draft.trim()) {
                const clash = sheets.some((x) => x.id !== s.id && x.name.toLocaleLowerCase('ru') === draft.trim().toLocaleLowerCase('ru'));
                if (clash) useUI.getState().toast({ text: 'Лист с таким названием уже есть', tone: 'error' });
                else store.renameSheet(s, draft);
              }
              setRenaming(null);
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
              if (e.key === 'Escape') setRenaming(null);
            }}
            size={Math.max(4, draft.length + 1)}
          />
        ) : (
          <button
            key={s.id}
            type="button"
            role="tab"
            aria-selected={s.id === active}
            className="tab"
            onClick={() => switchTo(s.id)}
            onDoubleClick={() => {
              setDraft(s.name);
              setRenaming(s.id);
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              set({ menu: { kind: 'sheet', x: e.clientX, y: e.clientY, sheetId: s.id } });
            }}
          >
            {s.name}
          </button>
        ),
      )}
      <button
        type="button"
        className="tab tab-add"
        aria-label="Добавить лист"
        title="Добавить лист"
        onClick={() => {
          let k = sheets.length + 1;
          while (sheets.some((x) => x.name === `Лист ${k}`)) k++;
          const sh = emptySheet(`Лист ${k}`, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'], 60);
          store.addSheet(sh);
          set({ sel: { ar: 0, ac: 0, fr: 0, fc: 0 } });
        }}
      >
        <Plus size={14} strokeWidth={2} />
      </button>
    </div>
  );
}

/** Сумма/среднее/количество по выделению — как в строке состояния Excel. */
function Aggregates() {
  const v = useStoreVersion();
  const sel = useUI((s) => s.sel);
  const stats = useMemo(() => {
    const { r1, r2, c1, c2 } = selRect(sel);
    const cells = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (cells < 2) return null;
    const { sheet } = ctx();
    let sum = 0;
    let nums = 0;
    let filled = 0;
    let nf = undefined as ReturnType<typeof store.display>['style']['nf'];
    const limit = 400000;
    let seen = 0;
    for (let vr = r1; vr <= r2 && seen < limit; vr++)
      for (let vc = c1; vc <= c2; vc++, seen++) {
        const x = cellAt(vr, vc);
        if (!x) continue;
        const val = store.cellValue(sheet.id, x.phys, x.c);
        if (val !== null && val !== '') filled++;
        if (typeof val === 'number') {
          sum += val;
          nums++;
          nf ??= store.styleOf(sheet, x.col, x.row?.cells[x.colId]).nf;
        }
      }
    return { sum, nums, filled, nf };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel, v]);
  if (!stats) return null;
  const fmt = (n: number) => formatNumber(parseFloat(n.toPrecision(12)), stats.nf && stats.nf.k !== 'date' ? stats.nf : { k: 'number', d: Number.isInteger(n) ? 0 : 2 });
  return (
    <div className="agg" aria-live="polite">
      {stats.nums > 0 && (
        <>
          <span>
            <span className="agg-k">Сумма</span> <span className="agg-v">{fmt(stats.sum)}</span>
          </span>
          <span>
            <span className="agg-k">Среднее</span> <span className="agg-v">{fmt(stats.sum / stats.nums)}</span>
          </span>
        </>
      )}
      <span>
        <span className="agg-k">Заполнено</span> <span className="agg-v">{stats.filled.toLocaleString('ru')}</span>
      </span>
    </div>
  );
}

function SaveIndicator() {
  const s = useSaveState();
  const text = s.status === 'saving' ? 'Сохраняем…' : s.status === 'error' ? 'Не сохранено — повторим' : 'Сохранено в браузере';
  return (
    <span className={'save save--' + s.status} role="status" title={s.error}>
      <span className="save-dot" aria-hidden />
      <span className="save-text">{text}</span>
    </span>
  );
}

export function StatusBar() {
  useStoreVersion();
  const { sheet, view } = ctx();
  const filters = Object.keys(sheet.filters).length;
  const hiddenRows = sheet.rowOrder.reduce((n, id) => n + (sheet.rows.get(id)?.hidden ? 1 : 0), 0);
  return (
    <footer className="status">
      <SheetTabs />
      <div className="status-info">
        {view.filtered ? (
          <span>
            Показано <b>{view.rows.length.toLocaleString('ru')}</b> из {view.totalRows.toLocaleString('ru')}
            {filters > 0 && (
              <button
                type="button"
                className="link-btn"
                onClick={() => {
                  store.clearFilters(sheet);
                  gridApi.focus();
                }}
              >
                сбросить {filters > 1 ? `${filters} фильтра` : 'фильтр'}
              </button>
            )}
          </span>
        ) : (
          <span>
            {view.totalRows.toLocaleString('ru')} {rowsWord(view.totalRows)}
          </span>
        )}
        {hiddenRows > 0 && (
          <button type="button" className="link-btn" onClick={() => store.showAllRows(sheet)}>
            скрыто {hiddenRows} — показать
          </button>
        )}
      </div>
      <Aggregates />
      <SaveIndicator />
    </footer>
  );
}
