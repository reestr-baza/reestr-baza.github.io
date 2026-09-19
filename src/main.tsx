import '@fontsource-variable/golos-text/wght.css';
import '@fontsource-variable/martian-mono/standard.css';
import './styles/tokens.css';
import './styles/app.css';

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './app/App';
import { setInstances, wireSaveState } from './app/instance';
import { Store } from './model/store';
import { lastSnapshotTime, loadWorkbook, Persister, requestPersistence, takeSnapshot, writeWholeWorkbook } from './storage/persist';
import { buildDemo } from './storage/seed';

const SNAPSHOT_EVERY = 20 * 60 * 1000;

async function boot() {
  const root = document.getElementById('root')!;
  let data = await loadWorkbook();
  let fresh = false;
  if (!data) {
    data = await buildDemo(import.meta.env.BASE_URL);
    await writeWholeWorkbook(data.meta, data.sheets);
    fresh = true;
  }
  const store = new Store(data.meta, data.sheets);
  const persister = new Persister(store);
  setInstances(store, persister);
  wireSaveState(persister);

  createRoot(root).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
  document.getElementById('boot')?.remove();

  // снимки: при открытии (если давно не было) и раз в 20 минут, если были изменения
  if (!fresh && Date.now() - (await lastSnapshotTime()) > 6 * 3600 * 1000) void takeSnapshot(store, 'При открытии');
  let lastSnapVersion = store.version;
  window.setInterval(() => {
    if (store.version !== lastSnapVersion) {
      lastSnapVersion = store.version;
      void takeSnapshot(store, 'Автоматически');
    }
  }, SNAPSHOT_EVERY);

  // просим браузер не очищать базу при первой же правке
  const off = store.onCommit(() => {
    off();
    void requestPersistence();
  });
}

boot().catch((e) => {
  console.error(e);
  const el = document.getElementById('boot');
  if (el) {
    el.textContent = '';
    const p = document.createElement('p');
    p.textContent =
      'Не удалось открыть базу в этом браузере. Если открыт режим инкогнито — откройте страницу в обычном окне. Подробности: ' +
      (e instanceof Error ? e.message : String(e));
    el.appendChild(p);
  }
});
