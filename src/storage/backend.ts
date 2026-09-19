import type { Row, SheetMeta, WorkbookMeta } from '../model/types';

/**
 * Где живут данные.
 * local  — демо (GitHub Pages): всё в IndexedDB этого браузера.
 * server — рабочая версия: база и фото на сервере, вход по паролю.
 * Режим определяется при запуске: если рядом с сайтом отвечает /api/me — значит, есть сервер.
 */
export type Mode = 'local' | 'server';
export let mode: Mode = 'local';

const API = import.meta.env.BASE_URL + 'api/';
/** Отличаем свои сохранения от чужих (другая вкладка, другое устройство). */
export const clientId = Math.random().toString(36).slice(2);

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string,
  ) {
    super(message);
  }
}

export async function detectMode(): Promise<'local' | 'server' | 'login'> {
  // сборка для GitHub Pages: сервера там нет, не спрашиваем (иначе в консоли лишняя ошибка 404)
  if (import.meta.env.VITE_STORAGE === 'local') return 'local';
  try {
    const r = await fetch(API + 'me', { credentials: 'same-origin', headers: { accept: 'application/json' } });
    if (r.ok && (r.headers.get('content-type') ?? '').includes('application/json')) {
      const me = (await r.json()) as { login?: string | null };
      mode = 'server';
      return me.login ? 'server' : 'login';
    }
  } catch {
    /* сервера нет — работаем в браузере */
  }
  mode = 'local';
  return 'local';
}

async function api(path: string, init: RequestInit = {}): Promise<Response> {
  const r = await fetch(API + path, {
    credentials: 'same-origin',
    ...init,
    // свой заголовок на каждом запросе — защита от подделки запросов с чужих сайтов
    headers: { 'x-reestr': '1', 'x-client': clientId, ...(init.headers ?? {}) },
  });
  if (r.status === 401) {
    window.dispatchEvent(new Event('reestr:logged-out'));
    throw new ApiError(401, 'Сессия закончилась — войдите заново');
  }
  if (!r.ok) {
    const text = await r.text().catch(() => '');
    throw new ApiError(r.status, text || `Сервер ответил ${r.status}`);
  }
  return r;
}

const json = (body: unknown): RequestInit => ({ method: 'POST', body: JSON.stringify(body), headers: { 'content-type': 'application/json' } });

export async function login(user: string, password: string): Promise<void> {
  const r = await fetch(API + 'login', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'content-type': 'application/json', 'x-reestr': '1' },
    body: JSON.stringify({ login: user, password }),
  });
  if (r.status === 429) throw new ApiError(429, 'Слишком много попыток. Подождите 15 минут');
  if (!r.ok) throw new ApiError(r.status, 'Неверный логин или пароль');
}

export async function logout(): Promise<void> {
  await fetch(API + 'logout', { method: 'POST', credentials: 'same-origin', headers: { 'x-reestr': '1' } }).catch(() => {});
}

// ─── книга ───────────────────────────────────────────────────────────────────

export type SheetDump = SheetMeta & { rows: Row[] };
export interface RowChange {
  sheetId: string;
  rowId: string;
  row: Row | null;
}
export interface SaveBatch {
  meta?: WorkbookMeta;
  sheets?: SheetMeta[];
  removedSheets?: string[];
  rows?: RowChange[];
}

let knownImages: Promise<Set<string>> | null = null;

export const remote = {
  async load(): Promise<{ meta: WorkbookMeta; sheets: SheetDump[]; rev: number } | null> {
    const r = await api('workbook');
    if (r.status === 204) return null;
    return r.json();
  },
  async replace(meta: WorkbookMeta, sheets: SheetDump[]): Promise<number> {
    const r = await api('workbook', json({ meta, sheets }));
    return (await r.json()).rev;
  },
  async save(batch: SaveBatch): Promise<number> {
    const r = await api('save', json(batch));
    return (await r.json()).rev;
  },
  async rev(): Promise<{ rev: number; by: string | null }> {
    return (await api('rev')).json();
  },

  images: {
    /** Есть ли уже такое фото на сервере. Список берём один раз — импорт сотен фото не шлёт сотни проверок. */
    async has(id: string): Promise<boolean> {
      knownImages ??= remote.images.list().then((ids) => new Set(ids));
      try {
        return (await knownImages).has(id);
      } catch {
        knownImages = null;
        return false;
      }
    },
    async put(rec: { id: string; full: Blob; thumb: Blob; w: number; h: number; name?: string }): Promise<void> {
      const headers = { 'x-w': String(rec.w), 'x-h': String(rec.h), 'x-name': encodeURIComponent(rec.name ?? '') };
      await api(`images/${rec.id}/thumb`, { method: 'PUT', body: rec.thumb, headers: { ...headers, 'content-type': rec.thumb.type } });
      await api(`images/${rec.id}/full`, { method: 'PUT', body: rec.full, headers: { ...headers, 'content-type': rec.full.type } });
      void knownImages?.then((s) => s.add(rec.id)).catch(() => {});
    },
    async meta(id: string): Promise<{ w: number; h: number; name?: string; bytes: number; created: number } | null> {
      const r = await fetch(API + `images/${id}/meta`, { credentials: 'same-origin' });
      return r.ok ? r.json() : null;
    },
    url(id: string, variant: 'thumb' | 'full'): string {
      return API + `images/${id}/${variant}`;
    },
    async blob(id: string, variant: 'thumb' | 'full'): Promise<Blob | null> {
      const r = await fetch(API + `images/${id}/${variant}`, { credentials: 'same-origin' });
      return r.ok ? r.blob() : null;
    },
    async list(): Promise<string[]> {
      return (await api('images')).json();
    },
    async gc(keep: string[]): Promise<{ removed: number; freed: number }> {
      knownImages = null;
      return (await api('images/gc', json({ keep }))).json();
    },
  },

  snapshots: {
    async add(label: string, rows: number, data: Uint8Array): Promise<void> {
      await api('snapshots', {
        method: 'POST',
        body: data as BlobPart as BodyInit,
        headers: { 'content-type': 'application/octet-stream', 'x-label': encodeURIComponent(label), 'x-rows': String(rows) },
      });
    },
    async list(): Promise<{ id: number; ts: number; label: string; rows: number }[]> {
      return (await api('snapshots')).json();
    },
    async read(id: number): Promise<Uint8Array> {
      return new Uint8Array(await (await api(`snapshots/${id}`)).arrayBuffer());
    },
  },

  async usage(): Promise<{ usage: number; disk: number; free: number }> {
    return (await api('usage')).json();
  },
};
