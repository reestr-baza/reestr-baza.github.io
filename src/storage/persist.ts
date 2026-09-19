import { gunzipSync, gzipSync, strFromU8, strToU8 } from 'fflate';
import type { CommitInfo, Store } from '../model/store';
import type { Row, RowId, Sheet, SheetMeta, WorkbookMeta } from '../model/types';
import { clientId, mode, remote } from './backend';
import { db, type Snapshot, type StoredRow } from './db';

/** Версия базы на сервере, которую мы видели последней — чтобы заметить правки с другого устройства. */
let lastRev = 0;

export type SaveState = { status: 'saved' | 'saving' | 'error'; at: number; error?: string };

type SaveListener = (s: SaveState) => void;

/** Полный слепок книги в виде простого объекта — для снимков, бэкапов и экспорта. */
export interface WorkbookDump {
  format: 'reestr';
  version: 1;
  meta: WorkbookMeta;
  sheets: (SheetMeta & { rows: Row[] })[];
}

export function dumpWorkbook(store: Store): WorkbookDump {
  return {
    format: 'reestr',
    version: 1,
    meta: store.meta,
    sheets: store.sheetList().map((s) => ({
      ...sheetMetaOf(s),
      rows: s.rowOrder.map((id) => s.rows.get(id)!).filter(Boolean),
    })),
  };
}

export function sheetsFromDump(dump: WorkbookDump): Sheet[] {
  return dump.sheets.map(({ rows, ...meta }) => ({
    ...meta,
    filters: meta.filters ?? {},
    frozen: meta.frozen ?? 0,
    density: meta.density ?? 'S',
    rowOrder: meta.rowOrder ?? rows.map((r) => r.id),
    rows: new Map(rows.map((r) => [r.id, r])),
  }));
}

export function sheetMetaOf(s: Sheet): SheetMeta {
  return {
    id: s.id,
    name: s.name,
    columns: s.columns,
    rowOrder: s.rowOrder,
    keyColId: s.keyColId,
    frozen: s.frozen,
    density: s.density,
    filters: s.filters,
    merges: s.merges,
  };
}

export async function loadWorkbook(): Promise<{ meta: WorkbookMeta; sheets: Sheet[] } | null> {
  if (mode === 'server') {
    const data = await remote.load();
    if (!data || !data.sheets.length) return null;
    lastRev = data.rev;
    const sheets = sheetsFromDump({ format: 'reestr', version: 1, meta: data.meta, sheets: data.sheets });
    for (const s of sheets) {
      // строки, потерянные порядком (или наоборот), не должны ломать лист
      const inOrder = new Set(s.rowOrder);
      s.rowOrder = s.rowOrder.filter((rid) => s.rows.has(rid));
      for (const rid of s.rows.keys()) if (!inOrder.has(rid)) s.rowOrder.push(rid);
    }
    const ids = sheets.map((s) => s.id);
    return { meta: { ...data.meta, sheetIds: ids, activeSheet: ids.includes(data.meta.activeSheet) ? data.meta.activeSheet : ids[0] }, sheets };
  }
  const d = await db();
  const meta = await d.get('meta', 'wb');
  if (!meta) return null;
  const sheets: Sheet[] = [];
  for (const id of meta.sheetIds) {
    const sm = await d.get('sheets', id);
    if (!sm) continue;
    const stored = await d.getAllFromIndex('rows', 'bySheet', id);
    const rows = new Map<RowId, Row>();
    for (const r of stored) {
      const { sheetId: _s, ...row } = r;
      rows.set(row.id, row);
    }
    // строки, потерянные порядком (или наоборот), не должны ломать лист
    const order = sm.rowOrder.filter((rid) => rows.has(rid));
    for (const rid of rows.keys()) if (!order.includes(rid) && order.length < 200000) order.push(rid);
    sheets.push({ ...sm, rowOrder: order, rows });
  }
  if (!sheets.length) return null;
  const sheetIds = sheets.map((s) => s.id);
  return {
    meta: { ...meta, sheetIds, activeSheet: sheetIds.includes(meta.activeSheet) ? meta.activeSheet : sheetIds[0] },
    sheets,
  };
}

/** Полная запись книги (первый запуск, восстановление, импорт). */
export async function writeWholeWorkbook(meta: WorkbookMeta, sheets: Sheet[]) {
  if (mode === 'server') {
    lastRev = await remote.replace(
      meta,
      sheets.map((s) => ({ ...sheetMetaOf(s), rows: s.rowOrder.map((id) => s.rows.get(id)!).filter(Boolean) })),
    );
    return;
  }
  const d = await db();
  const tx = d.transaction(['meta', 'sheets', 'rows'], 'readwrite');
  await tx.objectStore('rows').clear();
  await tx.objectStore('sheets').clear();
  await tx.objectStore('meta').put(meta, 'wb');
  for (const s of sheets) {
    await tx.objectStore('sheets').put(sheetMetaOf(s));
    for (const id of s.rowOrder) {
      const row = s.rows.get(id);
      if (row) tx.objectStore('rows').put({ ...row, sheetId: s.id });
    }
  }
  await tx.done;
}

/** Отложенная запись изменений: копим «грязные» строки и листы, пишем пачкой. */
export class Persister {
  private dirtyRows = new Map<string, { sheetId: string; rowId: string }>();
  private dirtySheets = new Set<string>();
  private removedSheets = new Set<string>();
  private dirtyMeta = false;
  private timer: number | null = null;
  private firstDirtyAt = 0;
  private flushing: Promise<void> | null = null;
  private listeners = new Set<SaveListener>();
  state: SaveState = { status: 'saved', at: Date.now() };
  private channel: BroadcastChannel | null = null;
  readonly tabId = Math.random().toString(36).slice(2);

  constructor(private store: Store) {
    store.onCommit((c) => this.onCommit(c));
    if ('BroadcastChannel' in window) this.channel = new BroadcastChannel('reestr');
    window.addEventListener('pagehide', () => void this.flush());
    window.addEventListener('beforeunload', (e) => {
      if (mode === 'server' && (this.state.status !== 'saved' || this.dirtyRows.size)) {
        void this.flush();
        e.preventDefault();
      }
    });
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') void this.flush();
    });
  }

  onSaveState(fn: SaveListener) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  onRemoteChange(fn: () => void) {
    if (mode === 'server') {
      // с сервером следим за версией базы: раз в 20 секунд и при возврате на вкладку
      const check = async () => {
        if (this.dirtyRows.size || this.flushing) return;
        try {
          const r = await remote.rev();
          if (r.rev > lastRev && r.by !== clientId) fn();
        } catch {
          /* нет связи — проверим позже */
        }
      };
      const timer = window.setInterval(check, 20000);
      const onFocus = () => void check();
      window.addEventListener('focus', onFocus);
      return () => {
        window.clearInterval(timer);
        window.removeEventListener('focus', onFocus);
      };
    }
    if (!this.channel) return () => {};
    const h = (e: MessageEvent) => {
      if (e.data?.type === 'changed' && e.data.from !== this.tabId) fn();
    };
    this.channel.addEventListener('message', h);
    return () => this.channel?.removeEventListener('message', h);
  }

  private setState(s: SaveState) {
    this.state = s;
    for (const l of this.listeners) l(s);
  }

  private onCommit(c: CommitInfo) {
    for (const r of c.rows) this.dirtyRows.set(r.sheetId + '/' + r.rowId, { sheetId: r.sheetId, rowId: r.rowId });
    for (const s of c.sheets) this.dirtySheets.add(s);
    for (const s of c.removedSheets) this.removedSheets.add(s);
    if (c.wb) this.dirtyMeta = true;
    this.schedule();
  }

  private schedule() {
    if (!this.firstDirtyAt) this.firstDirtyAt = Date.now();
    if (this.state.status !== 'saving') this.setState({ status: 'saving', at: Date.now() });
    if (this.timer) clearTimeout(this.timer);
    // не дольше 2 секунд копим изменения при непрерывной работе
    const wait = Date.now() - this.firstDirtyAt > 2000 ? 0 : 400;
    this.timer = window.setTimeout(() => void this.flush(), wait);
  }

  async flush(): Promise<void> {
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    if (this.flushing) await this.flushing;
    if (!this.dirtyRows.size && !this.dirtySheets.size && !this.dirtyMeta && !this.removedSheets.size) return;
    const rows = [...this.dirtyRows.values()];
    const sheets = [...this.dirtySheets];
    const removed = [...this.removedSheets];
    const meta = this.dirtyMeta;
    this.dirtyRows.clear();
    this.dirtySheets.clear();
    this.removedSheets.clear();
    this.dirtyMeta = false;
    this.firstDirtyAt = 0;

    this.flushing = (async () => {
      try {
        if (mode === 'server') {
          lastRev = await remote.save({
            meta: meta ? this.store.meta : undefined,
            sheets: sheets.map((id) => this.store.sheets.get(id)).filter((x): x is Sheet => !!x).map(sheetMetaOf),
            removedSheets: removed,
            rows: rows.map(({ sheetId, rowId }) => ({ sheetId, rowId, row: this.store.sheets.get(sheetId)?.rows.get(rowId) ?? null })),
          });
          this.setState({ status: 'saved', at: Date.now() });
          this.channel?.postMessage({ type: 'changed', from: this.tabId });
          return;
        }
        const d = await db();
        const tx = d.transaction(['meta', 'sheets', 'rows'], 'readwrite');
        const rowStore = tx.objectStore('rows');
        if (meta) void tx.objectStore('meta').put(this.store.meta, 'wb');
        for (const id of sheets) {
          const s = this.store.sheets.get(id);
          if (s) void tx.objectStore('sheets').put(sheetMetaOf(s));
        }
        for (const id of removed) {
          void tx.objectStore('sheets').delete(id);
          void rowStore.delete(IDBKeyRange.bound([id, ''], [id, '￿']));
        }
        for (const { sheetId, rowId } of rows) {
          const s = this.store.sheets.get(sheetId);
          const row = s?.rows.get(rowId);
          if (row) void rowStore.put({ ...row, sheetId } as StoredRow);
          else void rowStore.delete([sheetId, rowId]);
        }
        await tx.done;
        this.setState({ status: 'saved', at: Date.now() });
        this.channel?.postMessage({ type: 'changed', from: this.tabId });
      } catch (e) {
        console.error(e);
        // вернём изменения в очередь — попробуем ещё раз
        for (const r of rows) this.dirtyRows.set(r.sheetId + '/' + r.rowId, r);
        for (const s of sheets) this.dirtySheets.add(s);
        for (const s of removed) this.removedSheets.add(s);
        if (meta) this.dirtyMeta = true;
        this.setState({ status: 'error', at: Date.now(), error: e instanceof Error ? e.message : String(e) });
        this.timer = window.setTimeout(() => void this.flush(), 5000);
      }
    })();
    await this.flushing;
    this.flushing = null;
  }
}

// ─── автоматические снимки ───────────────────────────────────────────────────

const MAX_SNAPSHOTS = 15;

export function encodeDump(dump: WorkbookDump): Uint8Array {
  return gzipSync(strToU8(JSON.stringify(dump)), { level: 6 });
}

export function decodeDump(data: Uint8Array): WorkbookDump {
  return JSON.parse(strFromU8(gunzipSync(data)));
}

export async function takeSnapshot(store: Store, label: string): Promise<void> {
  const dump = dumpWorkbook(store);
  const rows = dump.sheets.reduce((n, s) => n + s.rows.length, 0);
  if (mode === 'server') return remote.snapshots.add(label, rows, encodeDump(dump));
  const d = await db();
  await d.add('snapshots', { ts: Date.now(), label, rows, data: encodeDump(dump) });
  const keys = await d.getAllKeys('snapshots');
  if (keys.length > MAX_SNAPSHOTS) {
    const tx = d.transaction('snapshots', 'readwrite');
    for (const k of keys.slice(0, keys.length - MAX_SNAPSHOTS)) void tx.store.delete(k);
    await tx.done;
  }
}

export async function listSnapshots(): Promise<Omit<Snapshot, 'data'>[]> {
  if (mode === 'server') return remote.snapshots.list();
  const d = await db();
  const all = await d.getAll('snapshots');
  return all.map(({ data: _d, ...rest }) => rest).reverse();
}

export async function readSnapshot(id: number): Promise<WorkbookDump | null> {
  if (mode === 'server') return decodeDump(await remote.snapshots.read(id));
  const d = await db();
  const s = await d.get('snapshots', id);
  return s ? decodeDump(s.data) : null;
}

export async function lastSnapshotTime(): Promise<number> {
  if (mode === 'server') return (await remote.snapshots.list())[0]?.ts ?? 0;
  const d = await db();
  const cursor = await d.transaction('snapshots').store.openCursor(null, 'prev');
  return cursor?.value.ts ?? 0;
}

/** Просим браузер не удалять данные при нехватке места. */
export async function requestPersistence(): Promise<boolean> {
  if (mode === 'server') return true;
  try {
    if (!navigator.storage?.persist) return false;
    if (await navigator.storage.persisted()) return true;
    return await navigator.storage.persist();
  } catch {
    return false;
  }
}

export async function storageInfo(): Promise<{ usage: number; quota: number; persisted: boolean }> {
  if (mode === 'server') {
    try {
      const u = await remote.usage();
      return { usage: u.usage, quota: u.disk, persisted: true };
    } catch {
      return { usage: 0, quota: 0, persisted: true };
    }
  }
  try {
    const est = await navigator.storage.estimate();
    const persisted = (await navigator.storage.persisted?.()) ?? false;
    return { usage: est.usage ?? 0, quota: est.quota ?? 0, persisted };
  } catch {
    return { usage: 0, quota: 0, persisted: false };
  }
}
