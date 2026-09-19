import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  ChevronDown,
  FilterX,
  ImagePlus,
  Italic,
  Link2,
  PaintBucket,
  Redo2,
  Strikethrough,
  TableCellsMerge,
  TableCellsSplit,
  Underline,
  Undo2,
  WrapText,
} from 'lucide-react';
import { useRef, useState, type ReactNode } from 'react';
import { currencyInfo } from '../model/format';
import type { Currency, Density, NumFmt } from '../model/types';
import { ColorPicker } from '../ui/ColorPicker';
import { RowsDensity, VAlignBottom, VAlignMiddle, VAlignTop } from '../ui/icons';
import { Menu, type MenuItem } from '../ui/Menu';
import { Popover } from '../ui/Popover';
import { useUI } from '../ui/state';
import { activeStyle, applyStyle, ctx, insertImages, mergeSelection, redo, selectionHasMerges, setNumFmt, toggleStyle, undo, unmergeSelection } from './actions';
import { gridApi } from './gridApi';
import { store, useStoreVersion } from './instance';
import { ParamsBar } from './ParamsBar';

const I = { size: 16, strokeWidth: 1.75 };

function Tb({
  label,
  hint,
  pressed,
  disabled,
  onClick,
  children,
  wide,
}: {
  label: string;
  hint?: string;
  pressed?: boolean;
  disabled?: boolean;
  onClick: (e: React.MouseEvent<HTMLButtonElement>) => void;
  children: ReactNode;
  wide?: boolean;
}) {
  return (
    <button
      type="button"
      className={'tb' + (wide ? ' tb--wide' : '')}
      aria-label={label}
      aria-pressed={pressed}
      disabled={disabled}
      data-tip={hint ? `${label} · ${hint}` : label}
      onMouseDown={(e) => e.preventDefault()}
      onClick={(e) => {
        onClick(e);
        gridApi.focus();
      }}
    >
      {children}
    </button>
  );
}

function fmtLabel(nf: NumFmt | undefined): string {
  switch (nf?.k) {
    case 'number':
      return 'Число';
    case 'currency':
      return currencyInfo(nf.c).label;
    case 'percent':
      return 'Процент';
    case 'date':
      return 'Дата';
    case 'text':
      return 'Текст';
    default:
      return 'Общий';
  }
}

export function Toolbar() {
  useStoreVersion();
  useUI((s) => s.sel);
  const set = useUI((s) => s.set);
  const st = activeStyle();
  const { sheet } = ctx();
  const [color, setColor] = useState<null | { kind: 'fg' | 'bg'; el: HTMLElement }>(null);
  const [fmtMenu, setFmtMenu] = useState<HTMLElement | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const nf = st.nf;
  const decimals = nf && 'd' in nf ? nf.d : nf?.k === 'general' || !nf ? null : 0;
  const hasFilters = Object.keys(sheet.filters).length > 0 || store.search !== '';
  const merged = selectionHasMerges();

  const cur = (c: Currency) => {
    const on = nf?.k === 'currency' && nf.c === c;
    return (
      <Tb key={c} label={`Формат: ${currencyInfo(c).label.toLowerCase()}`} pressed={on} onClick={() => setNumFmt(on ? undefined : { k: 'currency', c, d: c === 'RUB' ? 0 : c === 'CNY' ? 0 : 2 })}>
        <span className="tb-glyph">{currencyInfo(c).symbol}</span>
      </Tb>
    );
  };

  const changeDecimals = (delta: number) => {
    const base: NumFmt = !nf || nf.k === 'general' || nf.k === 'text' || nf.k === 'date' ? { k: 'number', d: 0 } : nf;
    if (!('d' in base)) return;
    setNumFmt({ ...base, d: Math.max(0, Math.min(6, base.d + delta)) } as NumFmt);
  };

  const fmtItems: MenuItem[] = [
    { label: 'Общий', hint: '31875,5', checked: !nf || nf.k === 'general', onSelect: () => setNumFmt(undefined) },
    { label: 'Число', hint: '31 875,50', checked: nf?.k === 'number', onSelect: () => setNumFmt({ k: 'number', d: 2 }) },
    'sep',
    { label: 'Рубли', hint: '31 875 ₽', checked: nf?.k === 'currency' && nf.c === 'RUB', onSelect: () => setNumFmt({ k: 'currency', c: 'RUB', d: 0 }) },
    { label: 'Юани', hint: '¥2 500', checked: nf?.k === 'currency' && nf.c === 'CNY', onSelect: () => setNumFmt({ k: 'currency', c: 'CNY', d: 0 }) },
    { label: 'Доллары', hint: '$1 250,00', checked: nf?.k === 'currency' && nf.c === 'USD', onSelect: () => setNumFmt({ k: 'currency', c: 'USD', d: 2 }) },
    { label: 'Евро', hint: '1 250,00 €', checked: nf?.k === 'currency' && nf.c === 'EUR', onSelect: () => setNumFmt({ k: 'currency', c: 'EUR', d: 2 }) },
    'sep',
    { label: 'Процент', hint: '15%', checked: nf?.k === 'percent', onSelect: () => setNumFmt({ k: 'percent', d: 0 }) },
    { label: 'Дата', hint: '18.09.2026', checked: nf?.k === 'date', onSelect: () => setNumFmt({ k: 'date' }) },
    { label: 'Текст', hint: 'как введено', checked: nf?.k === 'text', onSelect: () => setNumFmt({ k: 'text' }) },
  ];

  const density = (d: Density) => (
    <button
      key={d}
      type="button"
      className="seg-btn"
      aria-pressed={sheet.density === d}
      aria-label={{ S: 'Строки: компактно', M: 'Строки: средне', L: 'Строки: крупно, для фото' }[d]}
      data-tip={{ S: 'Компактные строки', M: 'Средние строки', L: 'Крупные строки для фото' }[d]}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => {
        store.allow(() => store.transact('Высота строк', () => store.setSheetMeta(sheet, { density: d })));
        gridApi.focus();
      }}
    >
      <RowsDensity level={d} />
    </button>
  );

  return (
    <div className="toolbar" role="toolbar" aria-label="Форматирование">
      <div className="tb-group">
        <Tb label="Отменить" hint="Ctrl+Z" disabled={!store.canUndo()} onClick={undo}>
          <Undo2 {...I} />
        </Tb>
        <Tb label="Повторить" hint="Ctrl+Y" disabled={!store.canRedo()} onClick={redo}>
          <Redo2 {...I} />
        </Tb>
      </div>

      <div className="tb-group">
        <button
          type="button"
          className="tb tb--select"
          aria-haspopup="menu"
          aria-expanded={!!fmtMenu}
          data-tip="Формат чисел"
          onMouseDown={(e) => e.preventDefault()}
          onClick={(e) => setFmtMenu(fmtMenu ? null : e.currentTarget)}
        >
          <span className="tb-select-text">{fmtLabel(nf)}</span>
          <ChevronDown size={14} strokeWidth={1.75} aria-hidden />
        </button>
        {cur('RUB')}
        {cur('CNY')}
        {cur('USD')}
        <Tb label="Формат: процент" pressed={nf?.k === 'percent'} onClick={() => setNumFmt(nf?.k === 'percent' ? undefined : { k: 'percent', d: 0 })}>
          <span className="tb-glyph">%</span>
        </Tb>
        <Tb label="Меньше знаков после запятой" disabled={decimals === 0} onClick={() => changeDecimals(-1)}>
          <span className="tb-glyph tb-glyph--dec">,0</span>
        </Tb>
        <Tb label="Больше знаков после запятой" onClick={() => changeDecimals(1)}>
          <span className="tb-glyph tb-glyph--dec">,00</span>
        </Tb>
      </div>

      <div className="tb-group">
        <Tb label="Жирный" hint="Ctrl+B" pressed={!!st.b} onClick={() => toggleStyle('b')}>
          <Bold {...I} strokeWidth={2.25} />
        </Tb>
        <Tb label="Курсив" hint="Ctrl+I" pressed={!!st.i} onClick={() => toggleStyle('i')}>
          <Italic {...I} />
        </Tb>
        <Tb label="Подчёркнутый" hint="Ctrl+U" pressed={!!st.u} onClick={() => toggleStyle('u')}>
          <Underline {...I} />
        </Tb>
        <Tb label="Зачёркнутый" hint="Ctrl+5" pressed={!!st.s} onClick={() => toggleStyle('s')}>
          <Strikethrough {...I} />
        </Tb>
        <Tb label="Цвет текста" onClick={(e) => setColor({ kind: 'fg', el: e.currentTarget })}>
          <span className="tb-color">
            <Baseline {...I} />
            <span className="tb-color-bar" style={{ background: st.fg ?? 'var(--ink)' }} />
          </span>
        </Tb>
        <Tb label="Заливка ячейки" onClick={(e) => setColor({ kind: 'bg', el: e.currentTarget })}>
          <span className="tb-color">
            <PaintBucket {...I} />
            <span className="tb-color-bar" style={{ background: st.bg ?? 'transparent', boxShadow: st.bg ? undefined : 'inset 0 0 0 1px var(--line-heavy)' }} />
          </span>
        </Tb>
      </div>

      <div className="tb-group">
        <Tb label="По левому краю" pressed={st.ha === 'left'} onClick={() => applyStyle({ ha: st.ha === 'left' ? undefined : 'left' }, 'Выравнивание')}>
          <AlignLeft {...I} />
        </Tb>
        <Tb label="По центру" pressed={st.ha === 'center'} onClick={() => applyStyle({ ha: st.ha === 'center' ? undefined : 'center' }, 'Выравнивание')}>
          <AlignCenter {...I} />
        </Tb>
        <Tb label="По правому краю" pressed={st.ha === 'right'} onClick={() => applyStyle({ ha: st.ha === 'right' ? undefined : 'right' }, 'Выравнивание')}>
          <AlignRight {...I} />
        </Tb>
        <span className="tb-gap" aria-hidden />
        <Tb label="По верху" pressed={st.va === 'top'} onClick={() => applyStyle({ va: st.va === 'top' ? undefined : 'top' }, 'Выравнивание')}>
          <VAlignTop />
        </Tb>
        <Tb label="По середине" pressed={!st.va || st.va === 'middle'} onClick={() => applyStyle({ va: undefined }, 'Выравнивание')}>
          <VAlignMiddle />
        </Tb>
        <Tb label="По низу" pressed={st.va === 'bottom'} onClick={() => applyStyle({ va: st.va === 'bottom' ? undefined : 'bottom' }, 'Выравнивание')}>
          <VAlignBottom />
        </Tb>
        <Tb label="Переносить текст" pressed={!!st.wrap} onClick={() => applyStyle({ wrap: !st.wrap || undefined }, 'Перенос текста')}>
          <WrapText {...I} />
        </Tb>
        {merged ? (
          <Tb label="Разъединить ячейки" pressed onClick={unmergeSelection}>
            <TableCellsSplit {...I} />
          </Tb>
        ) : (
          <Tb label="Объединить ячейки" hint="выделите несколько ячеек" onClick={mergeSelection}>
            <TableCellsMerge {...I} />
          </Tb>
        )}
      </div>

      <div className="tb-group">
        <Tb label="Ссылка" hint="Ctrl+K" onClick={() => set({ dialog: { kind: 'link', r: useUI.getState().sel.ar, c: useUI.getState().sel.ac } })}>
          <Link2 {...I} />
        </Tb>
        <Tb label="Фото в ячейку" hint="или перетащите файл" onClick={() => fileInput.current?.click()}>
          <ImagePlus {...I} />
        </Tb>
        <input
          ref={fileInput}
          type="file"
          accept="image/*"
          multiple
          hidden
          onChange={(e) => {
            const files = Array.from(e.target.files ?? []);
            e.target.value = '';
            const s = useUI.getState().sel;
            if (files.length) void insertImages(files, s.ar, s.ac);
          }}
        />
      </div>

      <div className="tb-group seg" role="group" aria-label="Высота строк">
        {density('S')}
        {density('M')}
        {density('L')}
      </div>

      {hasFilters && (
        <div className="tb-group tb-group--free">
          <button
            type="button"
            className="tb tb--text"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              store.clearFilters(sheet);
              store.setSearch('');
              useUI.getState().set({ searchOpen: false });
              gridApi.focus();
            }}
          >
            <FilterX {...I} />
            <span>Сбросить фильтры</span>
          </button>
        </div>
      )}

      <ParamsBar />

      {color && (
        <Popover anchor={{ el: color.el }} onClose={() => setColor(null)} className="pop--color" label={color.kind === 'fg' ? 'Цвет текста' : 'Заливка'}>
          <ColorPicker
            value={color.kind === 'fg' ? st.fg : st.bg}
            noneLabel={color.kind === 'fg' ? 'Обычный цвет текста' : 'Без заливки'}
            onPick={(c) => {
              applyStyle({ [color.kind]: c }, color.kind === 'fg' ? 'Цвет текста' : 'Заливка');
              setColor(null);
              gridApi.focus();
            }}
          />
        </Popover>
      )}
      {fmtMenu && <Menu anchor={{ el: fmtMenu }} items={fmtItems} onClose={() => setFmtMenu(null)} label="Формат чисел" />}
    </div>
  );
}
