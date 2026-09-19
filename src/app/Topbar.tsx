import { Archive, FileDown, FileSpreadsheet, FileUp, Keyboard, Menu as MenuIcon, Plus, Search, Settings2, X } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { BrandMark } from '../ui/icons';
import { Menu, type MenuItem } from '../ui/Menu';
import { useUI } from '../ui/state';
import { ctx, newCard, toast } from './actions';
import { gridApi } from './gridApi';
import { store, useStoreVersion } from './instance';

const I = { size: 16, strokeWidth: 1.75 };

function Title() {
  useStoreVersion();
  const [draft, setDraft] = useState<string | null>(null);
  const title = store.meta.title;
  return (
    <input
      className="wb-title"
      aria-label="Название базы"
      value={draft ?? title}
      spellCheck={false}
      onFocus={() => setDraft(title)}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => {
        const v = (draft ?? '').trim();
        if (v && v !== title) store.transact('Название', () => store.setWorkbookMeta({ title: v.slice(0, 80) }));
        setDraft(null);
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur();
        if (e.key === 'Escape') {
          setDraft(null);
          (e.target as HTMLInputElement).blur();
          gridApi.focus();
        }
      }}
      size={Math.max(6, (draft ?? title).length + 1)}
    />
  );
}

function SearchBox() {
  useStoreVersion();
  const open = useUI((s) => s.searchOpen);
  const set = useUI((s) => s.set);
  const [q, setQ] = useState(store.search);
  const input = useRef<HTMLInputElement>(null);
  const timer = useRef(0);

  useEffect(() => {
    if (open) input.current?.focus();
  }, [open]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.code === 'KeyF' && !document.querySelector('.dlg')) {
        e.preventDefault();
        set({ searchOpen: true });
        input.current?.focus();
        input.current?.select();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [set]);

  const apply = (v: string) => {
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => {
      store.setSearch(v);
      useUI.getState().setSel({ ar: 0, ac: 0, fr: 0, fc: 0 });
    }, 160);
  };

  const { sheet } = ctx();
  const view = store.view(sheet);
  const active = store.search.trim() !== '';
  return (
    <div className={'search' + (open || active ? ' is-open' : '')}>
      <Search {...I} className="search-ic" aria-hidden />
      <input
        ref={input}
        type="search"
        className="search-input"
        placeholder="Найти по всем столбцам"
        aria-label="Поиск по листу"
        value={q}
        onFocus={() => set({ searchOpen: true })}
        onBlur={() => !q && set({ searchOpen: false })}
        onChange={(e) => {
          setQ(e.target.value);
          apply(e.target.value);
        }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            setQ('');
            store.setSearch('');
            set({ searchOpen: false });
            gridApi.focus();
          }
          if (e.key === 'Enter') gridApi.focus();
        }}
      />
      {active && (
        <span className="search-count" role="status">
          {view.rows.length.toLocaleString('ru')}
        </span>
      )}
      {(q || open) && (
        <button
          type="button"
          className="icon-btn icon-btn--sm"
          aria-label="Очистить поиск"
          onClick={() => {
            setQ('');
            store.setSearch('');
            set({ searchOpen: false });
            gridApi.focus();
          }}
        >
          <X size={14} strokeWidth={1.75} />
        </button>
      )}
    </div>
  );
}

export function Topbar() {
  const set = useUI((s) => s.set);
  const [menu, setMenu] = useState<HTMLElement | null>(null);

  const items: MenuItem[] = [
    { label: 'Импорт из Excel или CSV…', icon: <FileUp {...I} />, onSelect: () => set({ dialog: { kind: 'import' } }) },
    {
      label: 'Скачать лист в Excel (.xlsx)',
      icon: <FileSpreadsheet {...I} />,
      onSelect: async () => {
        toast('Готовим файл Excel…');
        const { exportXlsx } = await import('../storage/xlsx');
        await exportXlsx(store, 'sheet');
      },
    },
    {
      label: 'Скачать всю базу в Excel',
      icon: <FileSpreadsheet {...I} />,
      onSelect: async () => {
        toast('Готовим файл Excel…');
        const { exportXlsx } = await import('../storage/xlsx');
        await exportXlsx(store, 'all');
      },
    },
    {
      label: 'Скачать лист в CSV',
      icon: <FileDown {...I} />,
      onSelect: async () => {
        const { exportCsv } = await import('../storage/xlsx');
        exportCsv(store);
      },
    },
    'sep',
    { label: 'Резервные копии…', icon: <Archive {...I} />, onSelect: () => set({ dialog: { kind: 'backup' } }) },
    { label: 'Параметры и курсы валют…', icon: <Settings2 {...I} />, onSelect: () => set({ dialog: { kind: 'settings' } }) },
    { label: 'Горячие клавиши', icon: <Keyboard {...I} />, hint: 'F1', onSelect: () => set({ dialog: { kind: 'shortcuts' } }) },
  ];

  return (
    <header className="topbar">
      <div className="brand">
        <BrandMark />
        <span className="brand-name">Реестр</span>
      </div>
      <span className="topbar-sep" aria-hidden />
      <Title />
      <div className="topbar-spacer" />
      <SearchBox />
      <button type="button" className="btn btn--primary topbar-new" onClick={newCard}>
        <Plus size={16} strokeWidth={2} aria-hidden />
        <span>Новая карточка</span>
      </button>
      <button
        type="button"
        className="btn btn--ghost topbar-file"
        aria-haspopup="menu"
        aria-expanded={!!menu}
        onClick={(e) => setMenu(menu ? null : e.currentTarget)}
      >
        <MenuIcon {...I} aria-hidden />
        <span>Файл</span>
      </button>
      {menu && <Menu anchor={{ el: menu }} items={items} onClose={() => setMenu(null)} placement="bottom-end" label="Файл" />}
    </header>
  );
}
