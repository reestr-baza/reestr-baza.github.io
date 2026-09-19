import { FileSpreadsheet } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { toast } from '../app/actions';
import { gridApi } from '../app/gridApi';
import { persister, store } from '../app/instance';
import { rowsWord } from '../app/StatusBar';
import type { NamedValue } from '../model/types';
import { takeSnapshot, writeWholeWorkbook } from '../storage/persist';
import type { ImportResult } from '../storage/xlsx';
import { Dialog } from '../ui/Dialog';
import { useUI } from '../ui/state';

function mergeNames(current: NamedValue[], incoming: NamedValue[]): NamedValue[] {
  const out = [...current];
  for (const n of incoming) {
    const i = out.findIndex((x) => x.name.toLocaleLowerCase('ru') === n.name.toLocaleLowerCase('ru'));
    if (i >= 0) out[i] = { ...out[i], value: n.value };
    else out.push(n);
  }
  return out;
}

export function ImportDialog({ initialFile }: { initialFile?: File | null }) {
  const set = useUI((s) => s.set);
  const [file, setFile] = useState<File | null>(initialFile ?? null);
  const [headerRows, setHeaderRows] = useState<'auto' | '0' | '1' | '2' | '3'>('auto');
  const [mode, setMode] = useState<'add' | 'replace'>(store.meta.demo ? 'replace' : 'add');
  const [progress, setProgress] = useState<string | null>(null);
  const [error, setError] = useState('');
  const [over, setOver] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => setFile(initialFile ?? null), [initialFile]);

  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };

  const run = async () => {
    if (!file) return;
    setError('');
    setProgress('Загружаем модуль Excel…');
    try {
      const mod = await import('../storage/xlsx');
      const opts = { headerRows: headerRows === 'auto' ? ('auto' as const) : Number(headerRows), onProgress: setProgress };
      const res: ImportResult = /\.csv$/i.test(file.name) ? await mod.importCsv(file, opts) : await mod.importXlsx(file, opts);
      await persister.flush();
      await takeSnapshot(store, `Перед импортом «${file.name}»`);
      if (mode === 'replace') {
        const meta = {
          ...store.meta,
          title: res.title || file.name.replace(/\.[^.]+$/, '').slice(0, 80) || store.meta.title,
          sheetIds: res.sheets.map((s) => s.id),
          activeSheet: res.sheets[0].id,
          names: res.names.length ? res.names : store.meta.demo ? [] : store.meta.names,
          demo: false,
        };
        await writeWholeWorkbook(meta, res.sheets);
        store.replaceAll(meta, res.sheets);
      } else {
        // имена листов не должны совпадать
        const used = new Set(store.sheetList().map((s) => s.name.toLocaleLowerCase('ru')));
        // импорт — осознанное действие через окно, его можно и в режиме просмотра
        store.allow(() => store.transact('Импорт', () => {
          for (const s of res.sheets) {
            let name = s.name;
            let k = 2;
            while (used.has(name.toLocaleLowerCase('ru'))) name = `${s.name} (${k++})`;
            used.add(name.toLocaleLowerCase('ru'));
            s.name = name;
            store.addSheet(s, s === res.sheets[0]);
          }
          if (res.names.length) store.setWorkbookMeta({ names: mergeNames(store.meta.names, res.names) });
        }));
      }
      set({ dialog: null, sel: { ar: 0, ac: 0, fr: 0, fc: 0 } });
      const params = res.names.length ? `. Из шапки взяты значения: ${res.names.map((n) => n.name).join(', ')}` : '';
      toast(
        `Импортировано: ${res.rows.toLocaleString('ru')} ${rowsWord(res.rows)}` + (res.images ? `, фото: ${res.images}` : '') + params,
      );
      // предупреждения (например, про маленькие фото) — отдельно и подольше, чтобы успели прочитать
      if (res.warnings.length) toast(res.warnings[0], { action: { label: 'Понятно', run: () => {} } });
      gridApi.focus();
    } catch (e) {
      console.error(e);
      setError(e instanceof Error ? `Не удалось прочитать файл: ${e.message}` : 'Не удалось прочитать файл');
      setProgress(null);
    }
  };

  return (
    <Dialog
      title="Импорт из Excel"
      onClose={close}
      className="dlg--md"
      footer={
        <>
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={close}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={run} disabled={!file || !!progress}>
            {progress ? 'Импортируем…' : 'Импортировать'}
          </button>
        </>
      }
    >
      <div
        className={'drop' + (over ? ' is-over' : '') + (file ? ' has-file' : '')}
        onDragOver={(e) => {
          e.preventDefault();
          setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          const f = e.dataTransfer.files[0];
          if (f) setFile(f);
        }}
      >
        <FileSpreadsheet size={28} strokeWidth={1.5} aria-hidden />
        {file ? (
          <p>
            <b>{file.name}</b> · {(file.size / 1024 / 1024).toFixed(1).replace('.', ',')} МБ
          </p>
        ) : (
          <p>Перетащите файл .xlsx или .csv сюда</p>
        )}
        <button type="button" className="btn btn--sm" onClick={() => input.current?.click()}>
          {file ? 'Выбрать другой' : 'Выбрать файл'}
        </button>
        <input
          ref={input}
          type="file"
          accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
          hidden
          onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = '';
            if (f) setFile(f);
          }}
        />
      </div>
      <p className="set-p set-p--muted">
        Переносятся значения, формулы, цвета, жирный и подчёркнутый шрифт, выравнивание, форматы ₽ и ¥, ширина столбцов, примечания и фото в ячейках. Значения из
        шапки, на которые ссылаются формулы (например, курс и доставка), станут параметрами — их видно над таблицей. Файлы .xls сначала сохраните в Excel как .xlsx.
      </p>
      <div className="form">
        <label className="form-row form-row--inline">
          <span className="form-label">Строк заголовка</span>
          <select className="input" value={headerRows} onChange={(e) => setHeaderRows(e.target.value as typeof headerRows)}>
            <option value="auto">Как закреплено в файле</option>
            <option value="1">1 строка</option>
            <option value="2">2 строки</option>
            <option value="3">3 строки</option>
            <option value="0">Нет заголовка</option>
          </select>
        </label>
        <fieldset className="radios">
          <legend className="form-label">Куда</legend>
          <label className="check">
            <input type="radio" name="imp-mode" checked={mode === 'replace'} onChange={() => setMode('replace')} />
            <span>{store.meta.demo ? 'Вместо примера' : 'Заменить всю базу'} (перед этим сохраним снимок)</span>
          </label>
          <label className="check">
            <input type="radio" name="imp-mode" checked={mode === 'add'} onChange={() => setMode('add')} />
            <span>Добавить новыми листами</span>
          </label>
        </fieldset>
      </div>
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
