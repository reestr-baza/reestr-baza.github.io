import { ArrowDownWideNarrow, ArrowUpNarrowWide, Search } from 'lucide-react';
import { useMemo, useState } from 'react';
import { ctx } from '../app/actions';
import { gridApi } from '../app/gridApi';
import { store } from '../app/instance';
import { compareScalars } from '../formula/coerce';
import type { FilterOp, FilterSpec } from '../model/types';
import { Popover } from '../ui/Popover';
import { useUI } from '../ui/state';

const EMPTY_KEY = '';
const MAX_LIST = 600;

const OPS: { op: FilterOp | ''; label: string; args: 0 | 1 | 2; numeric?: boolean }[] = [
  { op: '', label: 'Без условия', args: 0 },
  { op: 'contains', label: 'Содержит', args: 1 },
  { op: 'notContains', label: 'Не содержит', args: 1 },
  { op: 'begins', label: 'Начинается с', args: 1 },
  { op: 'ends', label: 'Заканчивается на', args: 1 },
  { op: 'eq', label: 'Равно', args: 1 },
  { op: 'neq', label: 'Не равно', args: 1 },
  { op: 'gt', label: 'Больше', args: 1, numeric: true },
  { op: 'gte', label: 'Больше или равно', args: 1, numeric: true },
  { op: 'lt', label: 'Меньше', args: 1, numeric: true },
  { op: 'lte', label: 'Меньше или равно', args: 1, numeric: true },
  { op: 'between', label: 'Между', args: 2, numeric: true },
  { op: 'top', label: 'Наибольшие N', args: 1, numeric: true },
  { op: 'empty', label: 'Пустые', args: 0 },
  { op: 'notEmpty', label: 'Непустые', args: 0 },
];

export function FilterMenu() {
  const fm = useUI((s) => s.filterMenu);
  const set = useUI((s) => s.set);
  if (!fm) return null;
  return <FilterPanel key={fm.c} vc={fm.c} rect={fm.rect} onClose={() => set({ filterMenu: null })} />;
}

function FilterPanel({ vc, rect, onClose }: { vc: number; rect: { left: number; top: number; bottom: number; right: number }; onClose: () => void }) {
  const { sheet, view } = ctx();
  const c = view.cols[vc];
  const col = sheet.columns[c];
  const spec: FilterSpec = sheet.filters[col.id] ?? {};

  // различные значения столбца: текст как в ячейке + количество
  const { values, colors, hasPhotos, isNumeric } = useMemo(() => {
    const rows = spec.values || spec.cond || spec.color || spec.photo ? sheet.rowOrder : view.rows;
    const counts = new Map<string, { n: number; v: unknown }>();
    const bg = new Map<string, number>();
    const fg = new Map<string, number>();
    let photos = 0;
    let nums = 0;
    let filled = 0;
    for (const id of rows) {
      const i = view.rowIndex.get(id)!;
      const d = store.display(sheet, i, c);
      const e = counts.get(d.text);
      if (e) e.n++;
      else counts.set(d.text, { n: 1, v: d.value });
      if (d.style.bg) bg.set(d.style.bg, (bg.get(d.style.bg) ?? 0) + 1);
      if (d.style.fg) fg.set(d.style.fg, (fg.get(d.style.fg) ?? 0) + 1);
      if (d.img) photos++;
      if (d.text) filled++;
      if (d.isNumber) nums++;
    }
    const list = [...counts.entries()].map(([k, x]) => ({ key: k, n: x.n, v: x.v }));
    list.sort((a, b) => {
      if (a.key === EMPTY_KEY) return 1;
      if (b.key === EMPTY_KEY) return -1;
      return compareScalars(a.v as never, b.v as never);
    });
    return {
      values: list,
      colors: { bg: [...bg.entries()], fg: [...fg.entries()] },
      hasPhotos: photos > 0,
      isNumeric: nums > 0 && nums >= filled * 0.6,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [col.id]);

  const [q, setQ] = useState('');
  const [allowed, setAllowed] = useState<Set<string>>(() => new Set(spec.values ?? values.map((v) => v.key)));
  const [op, setOp] = useState<FilterOp | ''>(spec.cond?.op ?? '');
  const [a, setA] = useState(spec.cond?.a ?? '');
  const [b, setB] = useState(spec.cond?.b ?? '');
  const [color, setColor] = useState<FilterSpec['color']>(spec.color);
  const [photo, setPhoto] = useState<FilterSpec['photo']>(spec.photo);

  const ql = q.trim().toLocaleLowerCase('ru');
  const shown = ql ? values.filter((v) => v.key.toLocaleLowerCase('ru').includes(ql)) : values;
  const allShownOn = shown.every((v) => allowed.has(v.key));

  const toggleAll = () => {
    const next = new Set(allowed);
    for (const v of shown) {
      if (allShownOn) next.delete(v.key);
      else next.add(v.key);
    }
    setAllowed(next);
  };

  const apply = () => {
    const everything = values.every((v) => allowed.has(v.key));
    // при поиске Excel оставляет только найденные значения
    const vals = ql ? shown.filter((v) => allowed.has(v.key)).map((v) => v.key) : [...allowed];
    const next: FilterSpec = {};
    if (ql || !everything) next.values = vals;
    const opDef = OPS.find((o) => o.op === op);
    if (op && opDef && (opDef.args === 0 || a.trim() !== '')) next.cond = { op, a: a.trim(), b: b.trim() };
    if (color) next.color = color;
    if (photo) next.photo = photo;
    store.setFilter(sheet, col.id, next);
    useUI.getState().setSel({ ar: 0, ac: vc, fr: 0, fc: vc });
    onClose();
    gridApi.focus();
  };

  const sort = (dir: 'asc' | 'desc') => {
    store.sortRows(sheet, c, dir);
    onClose();
    gridApi.focus();
  };

  const opDef = OPS.find((o) => o.op === op) ?? OPS[0];
  const ops = OPS.filter((o) => !o.numeric || isNumeric || o.op === op);

  return (
    <Popover anchor={{ rect }} onClose={onClose} className="pop--filter" role="dialog" label={`Фильтр: ${col.name}`} placement="bottom-start">
      <div className="flt" onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT' && apply()}>
        <div className="flt-title">{col.name}</div>
        <div className="flt-sort">
          <button type="button" className="flt-sort-btn" onClick={() => sort('asc')}>
            <ArrowUpNarrowWide size={15} strokeWidth={1.75} aria-hidden />
            {isNumeric ? 'От меньшего к большему' : 'От А до Я'}
          </button>
          <button type="button" className="flt-sort-btn" onClick={() => sort('desc')}>
            <ArrowDownWideNarrow size={15} strokeWidth={1.75} aria-hidden />
            {isNumeric ? 'От большего к меньшему' : 'От Я до А'}
          </button>
        </div>

        <div className="flt-sec">
          <label className="flt-label" htmlFor="flt-op">
            Условие
          </label>
          <div className="flt-cond">
            <select id="flt-op" className="input" value={op} onChange={(e) => setOp(e.target.value as FilterOp | '')}>
              {ops.map((o) => (
                <option key={o.op} value={o.op}>
                  {o.label}
                </option>
              ))}
            </select>
            {opDef.args >= 1 && (
              <input
                className="input"
                value={a}
                onChange={(e) => setA(e.target.value)}
                placeholder={opDef.op === 'top' ? '10' : opDef.numeric ? 'число' : 'текст'}
                inputMode={opDef.numeric ? 'decimal' : undefined}
                aria-label="Значение условия"
              />
            )}
            {opDef.args === 2 && (
              <input className="input" value={b} onChange={(e) => setB(e.target.value)} placeholder="и" inputMode="decimal" aria-label="Второе значение" />
            )}
          </div>
        </div>

        {(colors.bg.length > 0 || colors.fg.length > 0 || color) && (
          <div className="flt-sec">
            <span className="flt-label">По цвету</span>
            <div className="flt-colors">
              {colors.bg.map(([hex, n]) => (
                <button
                  key={'bg' + hex}
                  type="button"
                  className="flt-color"
                  aria-pressed={color?.kind === 'bg' && color.color === hex}
                  onClick={() => setColor(color?.kind === 'bg' && color.color === hex ? undefined : { kind: 'bg', color: hex })}
                  title={`Заливка ${hex} — ${n}`}
                >
                  <span className="flt-color-sw" style={{ background: hex }} aria-hidden />
                  <span className="flt-color-n">{n}</span>
                </button>
              ))}
              {colors.fg.map(([hex, n]) => (
                <button
                  key={'fg' + hex}
                  type="button"
                  className="flt-color"
                  aria-pressed={color?.kind === 'fg' && color.color === hex}
                  onClick={() => setColor(color?.kind === 'fg' && color.color === hex ? undefined : { kind: 'fg', color: hex })}
                  title={`Цвет текста ${hex} — ${n}`}
                >
                  <span className="flt-color-a" style={{ color: hex }} aria-hidden>
                    А
                  </span>
                  <span className="flt-color-n">{n}</span>
                </button>
              ))}
              {colors.bg.length > 0 && (
                <button
                  type="button"
                  className="flt-color"
                  aria-pressed={color?.kind === 'bg' && color.color === null}
                  onClick={() => setColor(color?.kind === 'bg' && color.color === null ? undefined : { kind: 'bg', color: null })}
                >
                  без заливки
                </button>
              )}
            </div>
          </div>
        )}

        {hasPhotos && (
          <div className="flt-sec">
            <span className="flt-label">Фото</span>
            <div className="seg seg--wide" role="radiogroup" aria-label="Фото">
              {([undefined, 'with', 'without'] as const).map((p) => (
                <button key={p ?? 'all'} type="button" role="radio" aria-checked={photo === p} className="seg-btn seg-btn--text" onClick={() => setPhoto(p)}>
                  {p === undefined ? 'Все' : p === 'with' ? 'С фото' : 'Без фото'}
                </button>
              ))}
            </div>
          </div>
        )}

        <div className="flt-sec flt-sec--values">
          <div className="flt-search">
            <Search size={14} strokeWidth={1.75} aria-hidden />
            <input className="input input--bare" placeholder="Поиск значений" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Поиск значений" />
          </div>
          <div className="flt-list" role="group" aria-label="Значения">
            <label className="flt-item flt-item--all">
              <input type="checkbox" checked={allShownOn} onChange={toggleAll} />
              <span>{ql ? 'Все найденные' : 'Выделить все'}</span>
              <span className="flt-n">{shown.reduce((s, v) => s + v.n, 0).toLocaleString('ru')}</span>
            </label>
            {shown.slice(0, MAX_LIST).map((v) => (
              <label key={v.key} className="flt-item">
                <input
                  type="checkbox"
                  checked={allowed.has(v.key)}
                  onChange={() => {
                    const next = new Set(allowed);
                    if (next.has(v.key)) next.delete(v.key);
                    else next.add(v.key);
                    setAllowed(next);
                  }}
                />
                <span className={v.key === EMPTY_KEY ? 'flt-empty' : ''}>{v.key === EMPTY_KEY ? '(пустые)' : v.key}</span>
                <span className="flt-n">{v.n.toLocaleString('ru')}</span>
              </label>
            ))}
            {shown.length > MAX_LIST && <div className="flt-more">Показаны первые {MAX_LIST} — уточните поиск</div>}
            {!shown.length && <div className="flt-more">Ничего не найдено</div>}
          </div>
        </div>

        <div className="flt-foot">
          <button
            type="button"
            className="btn btn--ghost"
            disabled={!sheet.filters[col.id]}
            onClick={() => {
              store.setFilter(sheet, col.id, null);
              onClose();
              gridApi.focus();
            }}
          >
            Сбросить
          </button>
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={onClose}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={apply}>
            Применить
          </button>
        </div>
      </div>
    </Popover>
  );
}
