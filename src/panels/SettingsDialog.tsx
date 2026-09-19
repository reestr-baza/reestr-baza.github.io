import { Plus, Trash2 } from 'lucide-react';
import { useState } from 'react';
import { toast } from '../app/actions';
import { requireEdit } from '../app/editMode';
import { gridApi } from '../app/gridApi';
import { store, useStoreVersion } from '../app/instance';
import { parseLooseNumber } from '../formula/coerce';
import { colToLetters } from '../formula/a1';
import type { NamedValue } from '../model/types';
import { buildStressSheet } from '../storage/seed';
import { Dialog } from '../ui/Dialog';
import { useUI } from '../ui/state';

const NAME_RE = /^[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_.]*$/;

export function SettingsDialog() {
  useStoreVersion();
  const set = useUI((s) => s.set);
  const sheet = store.activeSheet;
  const [names, setNames] = useState<{ name: string; value: string; note: string }[]>(() =>
    store.meta.names.map((n) => ({ name: n.name, value: String(n.value).replace('.', ','), note: n.note ?? '' })),
  );
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const close = () => {
    set({ dialog: null });
    gridApi.focus();
  };

  const saveNames = () => {
    const out: NamedValue[] = [];
    const seen = new Set<string>();
    for (const n of names) {
      const name = n.name.trim().toUpperCase();
      if (!name && !n.value.trim()) continue;
      if (!NAME_RE.test(name) || /^[A-Z]{1,3}\d+$/.test(name)) {
        setError(`«${n.name}» не подходит для имени: только буквы, цифры и _, без пробелов, не похоже на адрес ячейки`);
        return;
      }
      if (seen.has(name)) {
        setError(`Имя ${name} встречается дважды`);
        return;
      }
      seen.add(name);
      const num = parseLooseNumber(n.value);
      out.push({ name, value: num ?? n.value, note: n.note.trim() || undefined });
    }
    if (!requireEdit()) return;
    store.setNames(out);
    toast('Значения сохранены — формулы пересчитаны');
    close();
  };

  const stress = async () => {
    if (!requireEdit()) return;
    setBusy(true);
    await new Promise((r) => setTimeout(r, 30));
    const t0 = performance.now();
    const sh = buildStressSheet(20000);
    store.addSheet(sh);
    const t1 = performance.now();
    set({ sel: { ar: 0, ac: 0, fr: 0, fc: 0 }, dialog: null });
    setBusy(false);
    toast(`Лист на 20 000 строк создан за ${Math.round(t1 - t0)} мс. Попробуйте прокрутку, фильтр и сортировку`);
  };

  return (
    <Dialog
      title="Параметры"
      onClose={close}
      className="dlg--md"
      footer={
        <>
          <span className="flt-foot-gap" />
          <button type="button" className="btn" onClick={close}>
            Отмена
          </button>
          <button type="button" className="btn btn--primary" onClick={saveNames}>
            Сохранить
          </button>
        </>
      }
    >
      <section className="set-sec">
        <h3 className="set-h">Курсы и постоянные значения</h3>
        <p className="set-p">
          Меняете курс в одном месте — пересчитываются все формулы. Пример: <code>=G2*КУРС_ЮАНЯ</code>
        </p>
        <div className="nv">
          <div className="nv-head" aria-hidden>
            <span>Имя</span>
            <span>Значение</span>
            <span>Пояснение</span>
            <span />
          </div>
          {names.map((n, i) => (
            <div key={i} className="nv-row">
              <input
                className="input nv-name"
                aria-label="Имя"
                value={n.name}
                placeholder="КУРС_ЮАНЯ"
                onChange={(e) => {
                  const next = [...names];
                  next[i] = { ...n, name: e.target.value.toUpperCase().replace(/\s/g, '_') };
                  setNames(next);
                  setError('');
                }}
              />
              <input
                className="input input--num"
                aria-label="Значение"
                inputMode="decimal"
                value={n.value}
                onChange={(e) => {
                  const next = [...names];
                  next[i] = { ...n, value: e.target.value };
                  setNames(next);
                }}
              />
              <input
                className="input"
                aria-label="Пояснение"
                value={n.note}
                onChange={(e) => {
                  const next = [...names];
                  next[i] = { ...n, note: e.target.value };
                  setNames(next);
                }}
              />
              <button type="button" className="icon-btn" aria-label={`Удалить ${n.name}`} onClick={() => setNames(names.filter((_, k) => k !== i))}>
                <Trash2 size={15} strokeWidth={1.75} />
              </button>
            </div>
          ))}
          <button type="button" className="btn btn--ghost btn--sm" onClick={() => setNames([...names, { name: '', value: '', note: '' }])}>
            <Plus size={14} strokeWidth={2} aria-hidden /> Добавить значение
          </button>
        </div>
        {error && (
          <p className="form-error" role="alert">
            {error}
          </p>
        )}
      </section>

      <section className="set-sec">
        <h3 className="set-h">Карточка товара</h3>
        <label className="form-row form-row--inline">
          <span className="form-label">Открывать по клику на столбец</span>
          <select
            className="input"
            value={sheet.keyColId ?? ''}
            onChange={(e) => store.transact('Ключевой столбец', () => store.setSheetMeta(sheet, { keyColId: e.target.value || undefined }))}
          >
            <option value="">— только по номеру строки —</option>
            {sheet.columns.map((c, i) => (
              <option key={c.id} value={c.id}>
                {colToLetters(i)} · {c.name}
              </option>
            ))}
          </select>
        </label>
        <p className="set-p set-p--muted">Карточка также открывается кнопкой у номера строки и сочетанием Ctrl+Enter.</p>
      </section>

      <section className="set-sec">
        <h3 className="set-h">Проверка скорости</h3>
        <p className="set-p">Создаст отдельный лист на 20 000 строк с формулой столбца, чтобы проверить прокрутку, фильтры и сортировку на большом объёме.</p>
        <button type="button" className="btn" onClick={stress} disabled={busy}>
          {busy ? 'Создаём…' : 'Создать тестовый лист на 20 000 строк'}
        </button>
      </section>
    </Dialog>
  );
}
