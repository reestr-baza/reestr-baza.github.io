import { X } from 'lucide-react';
import { lazy, Suspense, useEffect, useState } from 'react';
import { ProductCard } from '../card/ProductCard';
import { FilterMenu } from '../grid/FilterMenu';
import { Grid } from '../grid/Grid';
import { LinkDialog, Lightbox, NoteDialog, RenameColumnDialog, ShortcutsDialog, Toasts } from '../panels/SmallDialogs';
import { useUI } from '../ui/state';
import { Tooltips } from '../ui/Tooltip';
import { ContextMenus } from './ContextMenus';
import { FormulaBar } from './FormulaBar';
import { persister, store, useStoreVersion } from './instance';
import { StatusBar } from './StatusBar';
import { Toolbar } from './Toolbar';
import { Topbar } from './Topbar';

const SettingsDialog = lazy(() => import('../panels/SettingsDialog').then((m) => ({ default: m.SettingsDialog })));
const BackupDialog = lazy(() => import('../panels/BackupDialog').then((m) => ({ default: m.BackupDialog })));
const ImportDialog = lazy(() => import('../panels/ImportDialog').then((m) => ({ default: m.ImportDialog })));

function Dialogs() {
  const dialog = useUI((s) => s.dialog);
  const [importFile, setImportFile] = useState<File | null>(null);

  useEffect(() => {
    const onFile = (e: Event) => {
      setImportFile((e as CustomEvent<File>).detail);
      useUI.getState().set({ dialog: { kind: 'import' } });
    };
    window.addEventListener('reestr:import-file', onFile);
    return () => window.removeEventListener('reestr:import-file', onFile);
  }, []);

  if (!dialog) return null;
  switch (dialog.kind) {
    case 'card':
      return <ProductCard rowId={dialog.rowId} />;
    case 'link':
      return <LinkDialog r={dialog.r} c={dialog.c} />;
    case 'rename-col':
      return <RenameColumnDialog c={dialog.c} />;
    case 'note':
      return <NoteDialog r={dialog.r} c={dialog.c} />;
    case 'shortcuts':
      return <ShortcutsDialog />;
    case 'settings':
      return (
        <Suspense fallback={null}>
          <SettingsDialog />
        </Suspense>
      );
    case 'backup':
      return (
        <Suspense fallback={null}>
          <BackupDialog />
        </Suspense>
      );
    case 'import':
      return (
        <Suspense fallback={null}>
          <ImportDialog initialFile={importFile} />
        </Suspense>
      );
  }
}

function DemoBanner() {
  useStoreVersion();
  const [hidden, setHidden] = useState(() => {
    try {
      return sessionStorage.getItem('reestr:demo-banner') === 'off';
    } catch {
      return false;
    }
  });
  if (!store.meta.demo || hidden) return null;
  return (
    <div className="banner banner--demo">
      <span>
        <b>Это пример на вымышленных данных.</b> Загрузите свой Excel — перенесутся строки, формулы, цвета, форматы ₽ и ¥ и фото из ячеек.
      </span>
      <button type="button" className="btn btn--sm btn--ink" onClick={() => useUI.getState().set({ dialog: { kind: 'import' } })}>
        Загрузить свой Excel
      </button>
      <button
        type="button"
        className="icon-btn icon-btn--sm"
        aria-label="Скрыть подсказку"
        onClick={() => {
          setHidden(true);
          try {
            sessionStorage.setItem('reestr:demo-banner', 'off');
          } catch {
            /* не критично */
          }
        }}
      >
        <X size={14} strokeWidth={1.75} />
      </button>
    </div>
  );
}

function RemoteChangeBanner() {
  const [stale, setStale] = useState(false);
  useEffect(() => persister.onRemoteChange(() => setStale(true)), []);
  if (!stale) return null;
  return (
    <div className="banner" role="alert">
      <span>База изменена в другой вкладке или на другом устройстве.</span>
      <button type="button" className="btn btn--sm" onClick={() => location.reload()}>
        Обновить
      </button>
    </div>
  );
}

export function App() {
  const editing = useUI((s) => s.editing);
  // файл, брошенный мимо клетки, браузер открыл бы вместо сайта — не даём
  useEffect(() => {
    const isFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files');
    const over = (e: DragEvent) => {
      if (isFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      if (!isFiles(e) || e.defaultPrevented) return;
      e.preventDefault();
      useUI.getState().toast({ text: 'Отпустите фото над ячейкой таблицы или клеткой карточки' });
    };
    window.addEventListener('dragover', over);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragover', over);
      window.removeEventListener('drop', drop);
    };
  }, []);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'F1') {
        e.preventDefault();
        useUI.getState().set({ dialog: { kind: 'shortcuts' } });
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  return (
    <div className={'app' + (editing ? '' : ' is-view')}>
      <a className="skip" href="#grid" onClick={(e) => {
        e.preventDefault();
        document.querySelector<HTMLElement>('.g')?.focus();
      }}>
        К таблице
      </a>
      <Topbar />
      <Toolbar />
      <FormulaBar />
      <DemoBanner />
      <RemoteChangeBanner />
      <main className="sheet-area" id="grid" aria-label={store.meta.title}>
        <Grid />
      </main>
      <StatusBar />
      <ContextMenus />
      <FilterMenu />
      <Dialogs />
      <Lightbox />
      <Toasts />
      <Tooltips />
    </div>
  );
}
