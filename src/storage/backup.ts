import { strFromU8, strToU8, unzip, zip, type Unzipped } from 'fflate';
import { downloadBlob, safeFileName, stamp } from '../lib/download';
import type { Store } from '../model/store';
import type { WorkbookDump } from './persist';
import { dumpWorkbook, sheetsFromDump, writeWholeWorkbook } from './persist';
import type { StoredImage } from './db';
import { getStoredImage, putStoredImage } from './images';

/** id всех фото, на которые ссылаются ячейки и блоки карточек. */
export function referencedImages(dump: WorkbookDump): Set<string> {
  const ids = new Set<string>();
  for (const s of dump.sheets)
    for (const r of s.rows) {
      for (const k in r.cells) {
        const img = r.cells[k].img;
        if (img) ids.add(img);
      }
      for (const b of r.card ?? []) if (b.img) ids.add(b.img);
    }
  return ids;
}

const ext = (type: string) => (type.includes('webp') ? 'webp' : type.includes('png') ? 'png' : 'jpg');

interface ImageManifest {
  id: string;
  w: number;
  h: number;
  name?: string;
  full: string;
  thumb: string;
  created: number;
}

/** Полная резервная копия: данные + все фото в одном .zip. */
export async function exportBackup(store: Store, onProgress?: (p: number) => void): Promise<number> {
  const dump = dumpWorkbook(store);
  const ids = [...referencedImages(dump)];
  const files: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
  const manifest: ImageManifest[] = [];
  let done = 0;
  for (const id of ids) {
    const rec = await getStoredImage(id);
    if (!rec) continue;
    const full = `images/${id}.${ext(rec.full.type)}`;
    const thumb = `thumbs/${id}.${ext(rec.thumb.type)}`;
    // фото уже сжаты — не тратим время на повторное сжатие
    files[full] = [new Uint8Array(await rec.full.arrayBuffer()), { level: 0 }];
    files[thumb] = [new Uint8Array(await rec.thumb.arrayBuffer()), { level: 0 }];
    manifest.push({ id, w: rec.w, h: rec.h, name: rec.name, full, thumb, created: rec.created });
    onProgress?.(++done / Math.max(1, ids.length));
  }
  files['reestr.json'] = strToU8(JSON.stringify({ ...dump, images: manifest }));
  const data = await new Promise<Uint8Array>((res, rej) => zip(files, { level: 6 }, (err, out) => (err ? rej(err) : res(out))));
  downloadBlob(new Blob([data as BlobPart], { type: 'application/zip' }), `${safeFileName(store.meta.title)}_копия_${stamp()}.zip`);
  return data.byteLength;
}

export interface BackupContents {
  dump: WorkbookDump;
  images: StoredImage[];
  rows: number;
}

export async function readBackup(file: File): Promise<BackupContents> {
  const buf = new Uint8Array(await file.arrayBuffer());
  const files = await new Promise<Unzipped>((res, rej) => unzip(buf, (err, out) => (err ? rej(err) : res(out))));
  const json = files['reestr.json'];
  if (!json) throw new Error('В архиве нет данных «Реестра» (reestr.json)');
  const parsed = JSON.parse(strFromU8(json)) as WorkbookDump & { images?: ImageManifest[] };
  if (parsed.format !== 'reestr') throw new Error('Файл не похож на резервную копию «Реестра»');
  const mime = (p: string) => (p.endsWith('webp') ? 'image/webp' : p.endsWith('png') ? 'image/png' : 'image/jpeg');
  const images: StoredImage[] = [];
  for (const m of parsed.images ?? []) {
    const f = files[m.full];
    const t = files[m.thumb];
    if (!f || !t) continue;
    const full = new Blob([f as BlobPart], { type: mime(m.full) });
    const thumb = new Blob([t as BlobPart], { type: mime(m.thumb) });
    images.push({ id: m.id, w: m.w, h: m.h, name: m.name, full, thumb, bytes: full.size + thumb.size, created: m.created });
  }
  const { images: _i, ...dump } = parsed;
  const rows = dump.sheets.reduce((n, s) => n + s.rows.length, 0);
  return { dump, images, rows };
}

/** Заменить текущую базу содержимым копии. */
export async function restoreDump(store: Store, dump: WorkbookDump, images: StoredImage[] = []) {
  for (const img of images) await putStoredImage(img);
  const sheets = sheetsFromDump(dump);
  const meta = { ...dump.meta, sheetIds: sheets.map((s) => s.id), activeSheet: sheets.some((s) => s.id === dump.meta.activeSheet) ? dump.meta.activeSheet : sheets[0].id };
  await writeWholeWorkbook(meta, sheets);
  store.replaceAll(meta, sheets);
}
