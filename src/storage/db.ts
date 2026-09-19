import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { Row, SheetMeta, WorkbookMeta } from '../model/types';

export interface StoredRow extends Row {
  sheetId: string;
}

export interface StoredImage {
  id: string;
  full: Blob;
  thumb: Blob;
  w: number;
  h: number;
  name?: string;
  bytes: number;
  created: number;
}

export interface Snapshot {
  id?: number;
  ts: number;
  label: string;
  rows: number;
  /** gzip(JSON) со всей книгой, без фото (фото хранятся отдельно и не удаляются, пока на них есть ссылки) */
  data: Uint8Array;
}

interface ReestrDB extends DBSchema {
  meta: { key: string; value: WorkbookMeta };
  sheets: { key: string; value: SheetMeta };
  rows: { key: [string, string]; value: StoredRow; indexes: { bySheet: string } };
  images: { key: string; value: StoredImage };
  snapshots: { key: number; value: Snapshot };
}

let dbPromise: Promise<IDBPDatabase<ReestrDB>> | null = null;

export function db() {
  dbPromise ??= openDB<ReestrDB>('reestr', 1, {
    upgrade(d) {
      d.createObjectStore('meta');
      d.createObjectStore('sheets', { keyPath: 'id' });
      const rows = d.createObjectStore('rows', { keyPath: ['sheetId', 'id'] });
      rows.createIndex('bySheet', 'sheetId');
      d.createObjectStore('images', { keyPath: 'id' });
      d.createObjectStore('snapshots', { keyPath: 'id', autoIncrement: true });
    },
    blocking() {
      // другая вкладка обновила схему — закрываемся, чтобы не мешать
      dbPromise?.then((x) => x.close());
      dbPromise = null;
    },
  });
  return dbPromise;
}

export type { ReestrDB };
