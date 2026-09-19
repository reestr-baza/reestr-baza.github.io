import { Download, RotateCcw, Upload } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { toast } from '../app/actions';
import { gridApi } from '../app/gridApi';
import { persister, store } from '../app/instance';
import { formatBytes } from '../lib/download';
import { exportBackup, readBackup, referencedImages, restoreDump, type BackupContents } from '../storage/backup';
import { collectGarbage } from '../storage/images';
import { mode } from '../storage/backend';
import { dumpWorkbook, listSnapshots, readSnapshot, requestPersistence, storageInfo, takeSnapshot } from '../storage/persist';
import { rowsWord } from '../app/StatusBar';
import { Dialog } from '../ui/Dialog';
import { useUI } from '../ui/state';

type Snap = Awaited<ReturnType<typeof listSnapshots>>[number];

const fmtTime = (ts: number) =>
  new Date(ts).toLocaleString('ru', { day: '2-digit', month: 'long', hour: '2-digit', minute: '2-digit' });

export function BackupDialog() {
  const set = useUI((s) => s.set);
  const [snaps, setSnaps] = useState<Snap[]>([]);
  const [info, setInfo] = useState<{ usage: number; quota: number; persisted: boolean } | null>(null);
  const [progress, setProgress] = useState<string | null>(null);
  const [pending, setPending] = useState<BackupContents | null>(null);
  const [error, setError] = useState('');
  const file = useRef<HTMLInputElement>(null);

  const refresh = async () => {
    setSnaps(await listSnapshots());
    setInfo(await storageInfo());
  };
  useEffect(() => {
    void refresh();
  }, []);

  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };

  const download = async () => {
    setError('');
    setProgress('Собираем архив…');
    try {
      await persister.flush();
      const bytes = await exportBackup(store, (p) => setProgress(`Упаковываем фото: ${Math.round(p * 100)}%`));
      localStorage.setItem('reestr:lastExport', String(Date.now()));
      toast(`Копия сохранена (${formatBytes(bytes)})`);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProgress(null);
    }
  };

  const restoreFrom = async (contents: BackupContents, label: string) => {
    setProgress('Восстанавливаем…');
    try {
      await persister.flush();
      await takeSnapshot(store, 'Перед восстановлением');
      await restoreDump(store, contents.dump, contents.images);
      set({ sel: { ar: 0, ac: 0, fr: 0, fc: 0 }, dialog: null });
      toast(`Восстановлено: ${label}`);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
      setProgress(null);
    }
  };

  const restoreSnapshot = async (s: Snap) => {
    const dump = await readSnapshot(s.id!);
    if (!dump) return;
    if (!confirm(`Вернуть базу к состоянию на ${fmtTime(s.ts)}? Текущее состояние сохранится отдельным снимком.`)) return;
    await restoreFrom({ dump, images: [], rows: s.rows }, `снимок от ${fmtTime(s.ts)}`);
  };

  const cleanup = async () => {
    setProgress('Ищем неиспользуемые фото…');
    const ref = referencedImages(dumpWorkbook(store));
    for (const s of await listSnapshots()) {
      const d = await readSnapshot(s.id!);
      if (d) referencedImages(d).forEach((id) => ref.add(id));
    }
    const res = await collectGarbage(ref);
    setProgress(null);
    toast(res.removed ? `Удалено фото: ${res.removed}, освобождено ${formatBytes(res.freed)}` : 'Неиспользуемых фото нет');
    void refresh();
  };

  const lastExport = Number(localStorage.getItem('reestr:lastExport') || 0);

  return (
    <Dialog title="Резервные копии" onClose={close} className="dlg--md">
      <section className="set-sec">
        <h3 className="set-h">Копия в файл</h3>
        <p className="set-p">
          Один архив .zip со всеми листами, формулами, оформлением и фото. Храните его на диске или в облаке — из него можно восстановить базу на любом
          компьютере.
        </p>
        <div className="set-actions">
          <button type="button" className="btn btn--primary" onClick={download} disabled={!!progress}>
            <Download size={15} strokeWidth={1.75} aria-hidden /> Скачать полную копию
          </button>
          <button type="button" className="btn" onClick={() => file.current?.click()} disabled={!!progress}>
            <Upload size={15} strokeWidth={1.75} aria-hidden /> Восстановить из файла
          </button>
          <input
            ref={file}
            type="file"
            accept=".zip,application/zip"
            hidden
            onChange={async (e) => {
              const f = e.target.files?.[0];
              e.target.value = '';
              if (!f) return;
              setError('');
              setProgress('Читаем архив…');
              try {
                setPending(await readBackup(f));
              } catch (err) {
                setError(err instanceof Error ? err.message : String(err));
              } finally {
                setProgress(null);
              }
            }}
          />
        </div>
        <p className="set-p set-p--muted">{lastExport ? `Последняя копия в файл: ${fmtTime(lastExport)}` : 'Копию в файл ещё не делали.'}</p>
        {pending && (
          <div className="confirm" role="alert">
            <p>
              В архиве: {pending.dump.sheets.length} {pending.dump.sheets.length === 1 ? 'лист' : 'листа'}, {pending.rows.toLocaleString('ru')} {rowsWord(pending.rows)}, фото:{' '}
              {pending.images.length}. Текущая база будет заменена — перед этим сохраним её снимок.
            </p>
            <div className="set-actions">
              <button type="button" className="btn btn--danger" onClick={() => restoreFrom(pending, 'копия из файла')}>
                Заменить базу
              </button>
              <button type="button" className="btn" onClick={() => setPending(null)}>
                Отмена
              </button>
            </div>
          </div>
        )}
      </section>

      <section className="set-sec">
        <div className="set-h-row">
          <h3 className="set-h">Автоматические снимки</h3>
          <button
            type="button"
            className="btn btn--ghost btn--sm"
            onClick={async () => {
              await takeSnapshot(store, 'Вручную');
              void refresh();
              toast('Снимок сохранён');
            }}
          >
            Сделать снимок сейчас
          </button>
        </div>
        <p className="set-p set-p--muted">Снимок делается при открытии и каждые 20 минут работы. Хранятся последние {mode === 'server' ? 30 : 15}.</p>
        {snaps.length ? (
          <ul className="snaps">
            {snaps.map((s) => (
              <li key={s.id} className="snap">
                <span className="snap-time">{fmtTime(s.ts)}</span>
                <span className="snap-meta">
                  {s.label} · {s.rows.toLocaleString('ru')} {rowsWord(s.rows)}
                </span>
                <button type="button" className="btn btn--ghost btn--sm" onClick={() => restoreSnapshot(s)}>
                  <RotateCcw size={14} strokeWidth={1.75} aria-hidden /> Вернуть
                </button>
              </li>
            ))}
          </ul>
        ) : (
          <p className="set-p">Снимков пока нет.</p>
        )}
      </section>

      <section className="set-sec">
        <h3 className="set-h">{mode === 'server' ? 'Место на сервере' : 'Хранилище браузера'}</h3>
        {info && mode === 'server' && (
          <p className="set-p">
            База и фото занимают {formatBytes(info.usage)}
            {info.quota ? ` — на диске сервера ${formatBytes(info.quota)}` : ''}. Каждую ночь сервер сам сохраняет копию базы и хранит последние 30.
          </p>
        )}
        {info && mode !== 'server' && (
          <p className="set-p">
            Занято {formatBytes(info.usage)}
            {info.quota ? ` из доступных ${formatBytes(info.quota)}` : ''}.{' '}
            {info.persisted ? (
              <span>Браузер не будет удалять эти данные при нехватке места.</span>
            ) : (
              <>
                <span>Браузер может очистить данные при нехватке места. </span>
                <button
                  type="button"
                  className="link-btn"
                  onClick={async () => {
                    const ok = await requestPersistence();
                    toast(ok ? 'Хранилище защищено от очистки' : 'Браузер не разрешил — делайте копии в файл');
                    void refresh();
                  }}
                >
                  Защитить
                </button>
              </>
            )}
          </p>
        )}
        <button type="button" className="btn btn--ghost btn--sm" onClick={cleanup} disabled={!!progress}>
          Удалить фото, которые нигде не используются
        </button>
      </section>
      {progress && (
        <p className="set-progress" role="status">
          {progress}
        </p>
      )}
      {error && (
        <p className="form-error" role="alert">
          {error}
        </p>
      )}
    </Dialog>
  );
}
