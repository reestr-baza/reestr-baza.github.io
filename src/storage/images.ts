import { mode, remote } from './backend';
import { db, type StoredImage } from './db';

const FULL_MAX = 2048;
const THUMB_MAX = 480;

let webpSupported: boolean | null = null;

async function encode(canvas: HTMLCanvasElement, quality: number): Promise<Blob> {
  if (webpSupported !== false) {
    const b = await new Promise<Blob | null>((res) => canvas.toBlob(res, 'image/webp', quality));
    if (b && b.type === 'image/webp') {
      webpSupported = true;
      return b;
    }
    webpSupported = false;
  }
  // Safari без WebP-кодировщика: JPEG на белом фоне
  const flat = document.createElement('canvas');
  flat.width = canvas.width;
  flat.height = canvas.height;
  const ctx = flat.getContext('2d')!;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, flat.width, flat.height);
  ctx.drawImage(canvas, 0, 0);
  const b = await new Promise<Blob | null>((res) => flat.toBlob(res, 'image/jpeg', quality));
  if (!b) throw new Error('Не удалось сжать изображение');
  return b;
}

function draw(src: CanvasImageSource, w: number, h: number, max: number): HTMLCanvasElement {
  const scale = Math.min(1, max / Math.max(w, h));
  const cw = Math.max(1, Math.round(w * scale));
  const ch = Math.max(1, Math.round(h * scale));
  const c = document.createElement('canvas');
  c.width = cw;
  c.height = ch;
  const ctx = c.getContext('2d')!;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(src, 0, 0, cw, ch);
  return c;
}

async function hashId(blob: Blob): Promise<string> {
  const buf = await blob.arrayBuffer();
  const digest = await crypto.subtle.digest('SHA-1', buf);
  return Array.from(new Uint8Array(digest).slice(0, 10), (b) => b.toString(16).padStart(2, '0')).join('');
}

async function decode(blob: Blob): Promise<{ src: CanvasImageSource; w: number; h: number; close(): void }> {
  if ('createImageBitmap' in window) {
    try {
      const bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' });
      return { src: bmp, w: bmp.width, h: bmp.height, close: () => bmp.close() };
    } catch {
      /* ниже — запасной путь через <img> */
    }
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.decoding = 'async';
    img.src = url;
    await img.decode();
    return { src: img, w: img.naturalWidth, h: img.naturalHeight, close: () => {} };
  } finally {
    URL.revokeObjectURL(url);
  }
}

export function isImageFile(f: { type: string; name?: string }): boolean {
  return /^image\/(png|jpe?g|webp|gif|bmp|avif|heic|heif)$/i.test(f.type) || /\.(png|jpe?g|webp|gif|bmp|avif)$/i.test(f.name ?? '');
}

async function exists(id: string): Promise<boolean> {
  if (mode === 'server') return remote.images.has(id);
  const d = await db();
  return !!(await d.getKey('images', id));
}

/**
 * Сжимает фото (до 2048px по большей стороне + превью 480px) и сохраняет:
 * в демо — в браузер, в рабочей версии — на сервер. Сжатие идёт до отправки, поэтому загрузка быстрая.
 * Одинаковые файлы хранятся один раз: id — хэш исходника.
 */
export async function importImage(blob: Blob, name?: string): Promise<string> {
  const id = await hashId(blob);
  if (await exists(id)) return id;
  const img = await decode(blob);
  try {
    const full = await encode(draw(img.src, img.w, img.h, FULL_MAX), 0.86);
    const thumb = await encode(draw(img.src, img.w, img.h, THUMB_MAX), 0.8);
    await putStoredImage({ id, full, thumb, w: img.w, h: img.h, name, bytes: full.size + thumb.size, created: Date.now() });
    return id;
  } finally {
    img.close();
  }
}

export async function putStoredImage(rec: StoredImage) {
  if (mode === 'server') return remote.images.put(rec);
  const d = await db();
  await d.put('images', rec);
}

/** Фото целиком (оба размера) — для выгрузки в Excel и резервной копии. */
export async function getStoredImage(id: string): Promise<StoredImage | undefined> {
  if (mode === 'server') {
    const [meta, full, thumb] = await Promise.all([remote.images.meta(id), remote.images.blob(id, 'full'), remote.images.blob(id, 'thumb')]);
    if (!meta || !full || !thumb) return undefined;
    return { id, full, thumb, w: meta.w, h: meta.h, name: meta.name, bytes: meta.bytes, created: meta.created };
  }
  const d = await db();
  return d.get('images', id);
}

/** Только размеры и имя — без загрузки самих картинок. */
export async function getImageMeta(id: string): Promise<{ w: number; h: number; name?: string } | null> {
  if (mode === 'server') return remote.images.meta(id);
  const rec = await (await db()).get('images', id);
  return rec ? { w: rec.w, h: rec.h, name: rec.name } : null;
}

// ─── адреса картинок ─────────────────────────────────────────────────────────

type Variant = 'thumb' | 'full';
const urls = new Map<string, string>();
const pending = new Map<string, Promise<string | null>>();
const LIMIT = 900;

export function peekImageUrl(id: string, variant: Variant): string | undefined {
  // с сервером адрес известен сразу, а кэш браузера сам держит картинки
  if (mode === 'server') return remote.images.url(id, variant);
  const key = variant + ':' + id;
  const u = urls.get(key);
  if (u) {
    // обновляем «свежесть» для LRU
    urls.delete(key);
    urls.set(key, u);
  }
  return u;
}

export function loadImageUrl(id: string, variant: Variant): Promise<string | null> {
  if (mode === 'server') return Promise.resolve(remote.images.url(id, variant));
  const key = variant + ':' + id;
  const have = urls.get(key);
  if (have) return Promise.resolve(have);
  let p = pending.get(key);
  if (!p) {
    p = (async () => {
      const rec = await getStoredImage(id);
      if (!rec) return null;
      const url = URL.createObjectURL(variant === 'thumb' ? rec.thumb : rec.full);
      urls.set(key, url);
      while (urls.size > LIMIT) {
        const [oldKey, oldUrl] = urls.entries().next().value!;
        urls.delete(oldKey);
        URL.revokeObjectURL(oldUrl);
      }
      return url;
    })().finally(() => pending.delete(key));
    pending.set(key, p);
  }
  return p;
}

/** Размер фото без загрузки картинки. */
export async function imageSize(id: string): Promise<{ w: number; h: number } | null> {
  const m = await getImageMeta(id);
  return m ? { w: m.w, h: m.h } : null;
}

/** Удаляет фото, на которые больше нет ссылок ни в данных, ни в снимках. */
export async function collectGarbage(referenced: Set<string>): Promise<{ removed: number; freed: number }> {
  if (mode === 'server') return remote.images.gc([...referenced]);
  const d = await db();
  const tx = d.transaction('images', 'readwrite');
  let cursor = await tx.store.openCursor();
  let removed = 0;
  let freed = 0;
  while (cursor) {
    if (!referenced.has(cursor.key)) {
      freed += cursor.value.bytes;
      removed++;
      await cursor.delete();
    }
    cursor = await cursor.continue();
  }
  await tx.done;
  return { removed, freed };
}

export async function allImageIds(): Promise<string[]> {
  if (mode === 'server') return remote.images.list();
  const d = await db();
  return d.getAllKeys('images');
}
