import { ChevronLeft, ChevronRight, Download, X } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { cellAt, ctx } from '../app/actions';
import { gridApi } from '../app/gridApi';
import { store } from '../app/instance';
import { colToLetters } from '../formula/a1';
import { isUrl, normalizeUrl, safeHref } from '../model/format';
import { getImageMeta } from '../storage/images';
import { Dialog } from '../ui/Dialog';
import { useUI } from '../ui/state';
import { useImageUrl } from '../ui/useImage';

export function LinkDialog({ r, c }: { r: number; c: number }) {
  const set = useUI((s) => s.set);
  const x = cellAt(r, c);
  const { sheet } = ctx();
  const cell = x?.row?.cells[x.colId];
  const currentText = cell?.v !== undefined ? String(cell.v) : '';
  const [text, setText] = useState(currentText);
  const [url, setUrl] = useState(cell?.href ?? (isUrl(currentText) ? currentText : ''));
  const [error, setError] = useState('');
  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };
  if (!x) return null;

  const save = () => {
    const u = url.trim();
    if (u && !safeHref(u)) {
      setError('Адрес должен начинаться с http:// или https:// (можно без него — добавим сами)');
      return;
    }
    store.transact(u ? 'Ссылка' : 'Удаление ссылки', () =>
      store.patchCell(sheet, x.rowId, x.colId, (cc) => {
        const next = { ...cc, f: undefined };
        next.href = u ? normalizeUrl(u) : undefined;
        next.v = text.trim() || (u ? u : undefined);
        return next;
      }),
    );
    close();
  };

  return (
    <Dialog
      title={`Ссылка в ячейке ${colToLetters(x.c)}${x.phys + 1}`}
      onClose={close}
      className="dlg--sm"
      footer={
        <>
          {cell?.href && (
            <button
              type="button"
              className="btn btn--ghost"
              onClick={() => {
                store.transact('Удаление ссылки', () => store.patchCell(sheet, x.rowId, x.colId, (cc) => ({ ...cc, href: undefined })));
                close();
              }}
            >
              Убрать ссылку
            </button>
          )}
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={close}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={save}>
            Сохранить
          </button>
        </>
      }
    >
      <form
        className="form"
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <label className="form-row">
          <span className="form-label">Адрес</span>
          <input
            className="input"
            data-autofocus
            inputMode="url"
            placeholder="https://detail.1688.com/…"
            value={url}
            aria-invalid={!!error}
            aria-describedby={error ? 'link-err' : undefined}
            onChange={(e) => {
              setUrl(e.target.value);
              setError('');
            }}
          />
        </label>
        {error && (
          <p id="link-err" className="form-error">
            {error}
          </p>
        )}
        <label className="form-row">
          <span className="form-label">Текст в ячейке</span>
          <input className="input" placeholder="Например: Поставщик на 1688" value={text} onChange={(e) => setText(e.target.value)} />
        </label>
        <button type="submit" hidden />
      </form>
    </Dialog>
  );
}

export function RenameColumnDialog({ c }: { c: number }) {
  const set = useUI((s) => s.set);
  const { sheet, view } = ctx();
  const col = sheet.columns[view.cols[c]];
  const [name, setName] = useState(col?.name ?? '');
  const [w, setW] = useState(String(col?.w ?? 120));
  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };
  if (!col) return null;
  const save = () => {
    const n = name.trim();
    const width = Math.max(36, Math.min(900, parseInt(w, 10) || col.w));
    store.updateColumn(sheet, col.id, { name: n || col.name, w: width }, 'Столбец');
    close();
  };
  return (
    <Dialog
      title={`Столбец ${colToLetters(view.cols[c])}`}
      onClose={close}
      className="dlg--sm"
      footer={
        <>
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={close}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={save}>
            Сохранить
          </button>
        </>
      }
    >
      <form
        className="form"
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <label className="form-row">
          <span className="form-label">Название</span>
          <input className="input" data-autofocus value={name} onChange={(e) => setName(e.target.value)} onFocus={(e) => e.target.select()} />
        </label>
        <label className="form-row">
          <span className="form-label">Ширина, px</span>
          <input className="input input--num" inputMode="numeric" value={w} onChange={(e) => setW(e.target.value.replace(/\D/g, ''))} />
        </label>
        <button type="submit" hidden />
      </form>
    </Dialog>
  );
}

export function NoteDialog({ r, c }: { r: number; c: number }) {
  const set = useUI((s) => s.set);
  const x = cellAt(r, c);
  const { sheet } = ctx();
  const [text, setText] = useState(x?.row?.cells[x.colId]?.note ?? '');
  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };
  if (!x) return null;
  const save = (value: string) => {
    store.transact(value ? 'Примечание' : 'Удаление примечания', () =>
      store.patchCell(sheet, x.rowId, x.colId, (cc) => ({ ...cc, note: value.trim() || undefined })),
    );
    close();
  };
  return (
    <Dialog
      title={`Примечание к ${colToLetters(x.c)}${x.phys + 1}`}
      onClose={close}
      className="dlg--sm"
      footer={
        <>
          {x.row?.cells[x.colId]?.note && (
            <button type="button" className="btn btn--ghost" onClick={() => save('')}>
              Удалить примечание
            </button>
          )}
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={close}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={() => save(text)}>
            Сохранить
          </button>
        </>
      }
    >
      <textarea
        className="input"
        rows={5}
        data-autofocus
        aria-label="Текст примечания"
        placeholder="Например: у продавца есть оригинал, спросить скидку"
        value={text}
        onChange={(e) => setText(e.target.value)}
        style={{ width: '100%' }}
      />
      <p className="set-p set-p--muted" style={{ marginTop: 8 }}>
        Ячейка с примечанием отмечена красным уголком, текст виден при наведении.
      </p>
    </Dialog>
  );
}

export function Lightbox() {
  const lb = useUI((s) => s.lightbox);
  const set = useUI((s) => s.set);
  const url = useImageUrl(lb?.imageId, 'full');
  const [meta, setMeta] = useState<{ w: number; h: number; name?: string } | null>(null);
  const list = lb?.list && lb.list.length > 1 ? lb.list : null;
  const index = list && lb ? list.indexOf(lb.imageId) : -1;
  const step = (d: number) => {
    if (!list || !lb) return;
    const next = list[(index + d + list.length) % list.length];
    set({ lightbox: { ...lb, imageId: next } });
  };
  const stepRef = useRef(step);
  stepRef.current = step;

  useEffect(() => {
    if (!lb) return;
    let alive = true;
    setMeta(null);
    void getImageMeta(lb.imageId).then((r) => alive && r && setMeta({ w: r.w, h: r.h, name: r.name }));
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        set({ lightbox: null });
      } else if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
        e.stopPropagation();
        e.preventDefault();
        stepRef.current(e.key === 'ArrowRight' ? 1 : -1);
      }
    };
    document.addEventListener('keydown', onKey, true);
    return () => {
      alive = false;
      document.removeEventListener('keydown', onKey, true);
    };
  }, [lb, set]);
  if (!lb) return null;
  const close = () => set({ lightbox: null });
  return createPortal(
    <div className="lb" role="dialog" aria-modal="true" aria-label="Просмотр фото" onPointerDown={(e) => e.target === e.currentTarget && close()}>
      <div className="lb-bar">
        <span className="lb-cap">
          {lb.caption}
          {meta && (
            <span className="lb-meta">
              {' '}
              · исходник {meta.w}×{meta.h}
            </span>
          )}
        </span>
        {url && (
          <a className="btn btn--ghost btn--on-dark" href={url} download={(meta?.name ?? 'foto').replace(/\.[a-z]+$/i, '') + '.webp'}>
            <Download size={15} strokeWidth={1.75} aria-hidden /> Скачать
          </a>
        )}
        <button type="button" className="icon-btn icon-btn--on-dark" onClick={close} aria-label="Закрыть просмотр" autoFocus>
          <X size={20} strokeWidth={1.75} />
        </button>
      </div>
      <div className="lb-stage">
        {list && (
          <button type="button" className="lb-nav lb-nav--prev" onClick={() => step(-1)} aria-label="Предыдущее фото">
            <ChevronLeft size={28} strokeWidth={1.5} />
          </button>
        )}
        {url ? <img className="lb-img" src={url} alt={lb.caption ?? 'Фото товара'} onClick={close} /> : <div className="lb-img lb-img--loading" />}
        {list && (
          <button type="button" className="lb-nav lb-nav--next" onClick={() => step(1)} aria-label="Следующее фото">
            <ChevronRight size={28} strokeWidth={1.5} />
          </button>
        )}
      </div>
      {list && (
        <div className="lb-count" aria-live="polite">
          {index + 1} из {list.length}
        </div>
      )}
    </div>,
    document.body,
  );
}

const SHORTCUTS: [string, string][] = [
  ['Стрелки', 'Перейти на соседнюю ячейку'],
  ['Shift + стрелки', 'Расширить выделение'],
  ['Ctrl + стрелки', 'К краю блока данных'],
  ['Enter / Tab', 'Сохранить и перейти вниз / вправо'],
  ['F2 или двойной клик', 'Редактировать ячейку'],
  ['Alt + Enter', 'Новая строка внутри ячейки'],
  ['Delete', 'Очистить выделенное'],
  ['Ctrl + C / X / V', 'Копировать / вырезать / вставить (и из Excel)'],
  ['Ctrl + Z / Y', 'Отменить / повторить'],
  ['Ctrl + B / I / U', 'Жирный / курсив / подчёркнутый'],
  ['Ctrl + D', 'Заполнить вниз'],
  ['Ctrl + K', 'Ссылка'],
  ['Ctrl + F', 'Поиск по листу'],
  ['Ctrl + Enter', 'Открыть карточку строки'],
  ['Ctrl + Пробел / Shift + Пробел', 'Выделить столбец / строку'],
  ['← / →  в карточке', 'Предыдущая / следующая карточка'],
];

export function ShortcutsDialog() {
  const set = useUI((s) => s.set);
  return (
    <Dialog title="Горячие клавиши" onClose={() => set({ dialog: null })} className="dlg--md">
      <dl className="keys">
        {SHORTCUTS.map(([k, v]) => (
          <div key={k} className="keys-row">
            <dt>
              {k.split(' / ').map((part, i) => (
                <span key={i}>
                  {i > 0 && ' / '}
                  <kbd>{part}</kbd>
                </span>
              ))}
            </dt>
            <dd>{v}</dd>
          </div>
        ))}
      </dl>
    </Dialog>
  );
}

export function Toasts() {
  const toasts = useUI((s) => s.toasts);
  const dismiss = useUI((s) => s.dismissToast);
  return createPortal(
    <div className="toasts" role="status" aria-live="polite">
      {toasts.map((t) => (
        <div key={t.id} className={'toast' + (t.tone === 'error' ? ' toast--error' : '')} role={t.tone === 'error' ? 'alert' : undefined}>
          <span className="toast-text">{t.text}</span>
          {t.action && (
            <button
              type="button"
              className="toast-action"
              onClick={() => {
                t.action!.run();
                dismiss(t.id);
              }}
            >
              {t.action.label}
            </button>
          )}
          <button type="button" className="toast-x" aria-label="Скрыть уведомление" onClick={() => dismiss(t.id)}>
            <X size={14} strokeWidth={1.75} />
          </button>
        </div>
      ))}
    </div>,
    document.body,
  );
}
