import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Baseline,
  Bold,
  ChevronLeft,
  ChevronRight,
  ImageMinus,
  ImagePlus,
  Italic,
  MoreHorizontal,
  PaintBucket,
  Plus,
  Printer,
  Trash2,
  Underline,
  X,
} from 'lucide-react';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { ctx, setSelection, toast } from '../app/actions';
import { gridApi } from '../app/gridApi';
import { store, useStoreVersion } from '../app/instance';
import type { CardBlock, CellStyle, Column, Row, Sheet } from '../model/types';
import { importImage, isImageFile } from '../storage/images';
import { ColorPicker } from '../ui/ColorPicker';
import { VAlignBottom, VAlignMiddle, VAlignTop } from '../ui/icons';
import { Menu } from '../ui/Menu';
import { Popover } from '../ui/Popover';
import { useUI } from '../ui/state';
import { useImageUrl } from '../ui/useImage';
import { barcodeBars } from './barcode';

const I = { size: 16, strokeWidth: 1.75 };
const BLOCKS_PER_ROW = 3;

function Barcode({ text, vertical, className }: { text: string; vertical?: boolean; className?: string }) {
  const bc = useMemo(() => barcodeBars(text), [text]);
  if (!bc) return null;
  const h = 40;
  return (
    <svg
      className={className}
      style={vertical ? { height: Math.min(340, (bc.modules + 20) * 3.2) } : undefined}
      viewBox={vertical ? `0 -10 ${h} ${bc.modules + 20}` : `-10 0 ${bc.modules + 20} ${h}`}
      preserveAspectRatio="none"
      role="img"
      aria-label={`Штрихкод ${text}`}
    >
      {bc.bars.map((b, i) =>
        vertical ? <rect key={i} x={0} y={b.x} width={h} height={b.w} /> : <rect key={i} x={b.x} y={0} width={b.w} height={h} />,
      )}
    </svg>
  );
}

function titleOf(sheet: Sheet, phys: number, keyC: number): string {
  // заголовок карточки — первые текстовые поля после артикула
  const parts: string[] = [];
  for (let c = 0; c < sheet.columns.length && parts.length < 4; c++) {
    if (c === keyC) continue;
    const d = store.display(sheet, phys, c);
    if (!d.text || d.img || d.isNumber || d.isError || d.href) continue;
    if (d.text.length > 40) continue;
    parts.push(d.text.replace(/\s*\n\s*/g, ' ').trim());
  }
  return parts.join(' · ');
}

// ─── блок карточки ───────────────────────────────────────────────────────────

/** Сильнее увеличивать маленькое превью бессмысленно — только размоется */
const MAX_UPSCALE = 1.6;

function BlockImage({ id, onOpen }: { id: string; onOpen: () => void }) {
  const url = useImageUrl(id, 'full');
  const [nat, setNat] = useState<{ w: number; h: number } | null>(null);
  const [box, setBox] = useState<{ w: number; h: number } | null>(null);
  const ref = useRef<HTMLButtonElement>(null);
  useEffect(() => setNat(null), [id]);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const ro = new ResizeObserver(([e]) => setBox({ w: e.contentRect.width, h: e.contentRect.height }));
    ro.observe(el);
    return () => ro.disconnect();
  }, [url]);
  if (!url) return <div className="blk-img blk-img--loading" aria-hidden />;
  // фото вписывается в блок целиком; если для этого его пришлось бы растянуть сильнее MAX_UPSCALE — оставляем меньше и помечаем
  const small = !!nat && !!box && Math.min(box.w / nat.w, box.h / nat.h) > MAX_UPSCALE;
  return (
    <button ref={ref} type="button" className="blk-img" onClick={onOpen} aria-label="Открыть фото крупно">
      <img
        src={url}
        alt=""
        draggable={false}
        onLoad={(e) => setNat({ w: e.currentTarget.naturalWidth, h: e.currentTarget.naturalHeight })}
        style={small ? { maxWidth: nat.w * MAX_UPSCALE, maxHeight: nat.h * MAX_UPSCALE } : undefined}
      />
      {small && (
        <span className="blk-lowres" title="Фото пришло из Excel маленьким превью. Для карточки загрузите оригинал — например, с телефона">
          маленькое фото {nat.w}×{nat.h}
        </span>
      )}
    </button>
  );
}

function AutoText({ value, onChange, style, placeholder, onFocus }: { value: string; onChange: (v: string) => void; style: React.CSSProperties; placeholder: string; onFocus: () => void }) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  useLayoutEffect(() => {
    const ta = ref.current;
    if (!ta) return;
    ta.style.height = '0px';
    ta.style.height = `${ta.scrollHeight}px`;
  }, [draft]);
  return (
    <textarea
      ref={ref}
      className="blk-text"
      style={style}
      value={draft}
      rows={1}
      placeholder={placeholder}
      aria-label="Текст блока"
      onFocus={onFocus}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => draft !== value && onChange(draft)}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          setDraft(value);
          const card = (e.target as HTMLElement).closest<HTMLElement>('.card');
          (e.target as HTMLTextAreaElement).blur();
          // следующий Escape закроет карточку — фокус не должен «выпасть» на страницу
          card?.focus({ preventScroll: true });
          e.stopPropagation();
        }
      }}
    />
  );
}

function Block({
  block,
  index,
  active,
  onActivate,
  onChange,
  onOpenImage,
}: {
  block: CardBlock;
  index: number;
  active: boolean;
  onActivate: () => void;
  onChange: (b: CardBlock, label: string) => void;
  onOpenImage: (id: string) => void;
}) {
  const [over, setOver] = useState(false);
  const st = block.st ?? {};
  const textStyle: React.CSSProperties = {
    textAlign: st.ha ?? 'left',
    fontWeight: st.b ? 600 : undefined,
    fontStyle: st.i ? 'italic' : undefined,
    textDecorationLine: st.u ? 'underline' : undefined,
    color: st.fg,
  };
  return (
    <div
      className={'blk' + (active ? ' is-active' : '') + (over ? ' is-drop' : '') + (block.img ? ' has-img' : '')}
      style={{ backgroundColor: st.bg }}
      data-va={st.va ?? 'bottom'}
      onPointerDown={onActivate}
      onFocus={onActivate}
      onDragOver={(e) => {
        if (Array.from(e.dataTransfer.types).includes('Files')) {
          e.preventDefault();
          setOver(true);
        }
      }}
      onDragLeave={() => setOver(false)}
      onDrop={() => setOver(false)}
      onPaste={async (e) => {
        const f = Array.from(e.clipboardData.files).find(isImageFile);
        if (!f) return;
        e.preventDefault();
        const id = await importImage(f, f.name);
        onChange({ ...block, img: id }, 'Фото в блоке');
      }}
      aria-label={`Блок ${index + 1}`}
      role="group"
      data-block={index}
    >
      {block.img && <BlockImage id={block.img} onOpen={() => onOpenImage(block.img!)} />}
      <div className="blk-textwrap">
        <AutoText
          value={block.text ?? ''}
          style={textStyle}
          placeholder={block.img ? 'Подпись' : 'Текст или фото'}
          onFocus={onActivate}
          onChange={(v) => onChange({ ...block, text: v || undefined }, 'Текст блока')}
        />
      </div>
      {!block.img && !block.text && <span className="blk-hint" aria-hidden>перетащите фото</span>}
    </div>
  );
}

function BlockToolbar({ block, onChange }: { block: CardBlock; onChange: (b: CardBlock, label: string) => void }) {
  const [color, setColor] = useState<null | { kind: 'fg' | 'bg'; el: HTMLElement }>(null);
  const file = useRef<HTMLInputElement>(null);
  const st = block.st ?? {};
  const patch = (p: Partial<CellStyle>, label: string) => {
    const next: CellStyle = { ...st, ...p };
    for (const k of Object.keys(next) as (keyof CellStyle)[]) if (next[k] === undefined || next[k] === false) delete next[k];
    onChange({ ...block, st: Object.keys(next).length ? next : undefined }, label);
  };
  const Btn = ({ label, on, onClick, children }: { label: string; on?: boolean; onClick: (e: React.MouseEvent<HTMLButtonElement>) => void; children: React.ReactNode }) => (
    <button type="button" className="tb" aria-label={label} data-tip={label} aria-pressed={on} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {children}
    </button>
  );
  return (
    <div className="blk-tools" role="toolbar" aria-label="Оформление блока">
      <Btn label={block.img ? 'Заменить фото' : 'Добавить фото'} onClick={() => file.current?.click()}>
        <ImagePlus {...I} />
      </Btn>
      {block.img && (
        <Btn label="Убрать фото" onClick={() => onChange({ ...block, img: undefined }, 'Убрать фото')}>
          <ImageMinus {...I} />
        </Btn>
      )}
      <span className="tb-sep" aria-hidden />
      <Btn label="По левому краю" on={!st.ha || st.ha === 'left'} onClick={() => patch({ ha: undefined }, 'Выравнивание')}>
        <AlignLeft {...I} />
      </Btn>
      <Btn label="По центру" on={st.ha === 'center'} onClick={() => patch({ ha: 'center' }, 'Выравнивание')}>
        <AlignCenter {...I} />
      </Btn>
      <Btn label="По правому краю" on={st.ha === 'right'} onClick={() => patch({ ha: 'right' }, 'Выравнивание')}>
        <AlignRight {...I} />
      </Btn>
      <span className="tb-sep" aria-hidden />
      <Btn label="По верху" on={st.va === 'top'} onClick={() => patch({ va: 'top' }, 'Выравнивание')}>
        <VAlignTop />
      </Btn>
      <Btn label="По середине" on={st.va === 'middle'} onClick={() => patch({ va: 'middle' }, 'Выравнивание')}>
        <VAlignMiddle />
      </Btn>
      <Btn label="По низу" on={!st.va || st.va === 'bottom'} onClick={() => patch({ va: undefined }, 'Выравнивание')}>
        <VAlignBottom />
      </Btn>
      <span className="tb-sep" aria-hidden />
      <Btn label="Жирный" on={!!st.b} onClick={() => patch({ b: !st.b }, 'Жирный')}>
        <Bold {...I} strokeWidth={2.25} />
      </Btn>
      <Btn label="Курсив" on={!!st.i} onClick={() => patch({ i: !st.i }, 'Курсив')}>
        <Italic {...I} />
      </Btn>
      <Btn label="Подчёркнутый" on={!!st.u} onClick={() => patch({ u: !st.u }, 'Подчёркнутый')}>
        <Underline {...I} />
      </Btn>
      <Btn label="Цвет текста" onClick={(e) => setColor({ kind: 'fg', el: e.currentTarget })}>
        <span className="tb-color">
          <Baseline {...I} />
          <span className="tb-color-bar" style={{ background: st.fg ?? 'var(--ink)' }} />
        </span>
      </Btn>
      <Btn label="Заливка" onClick={(e) => setColor({ kind: 'bg', el: e.currentTarget })}>
        <span className="tb-color">
          <PaintBucket {...I} />
          <span className="tb-color-bar" style={{ background: st.bg ?? 'transparent', boxShadow: st.bg ? undefined : 'inset 0 0 0 1px var(--line-heavy)' }} />
        </span>
      </Btn>
      <input
        ref={file}
        type="file"
        accept="image/*"
        hidden
        onChange={async (e) => {
          const f = e.target.files?.[0];
          e.target.value = '';
          if (!f) return;
          const id = await importImage(f, f.name);
          onChange({ ...block, img: id }, 'Фото в блоке');
        }}
      />
      {color && (
        <Popover anchor={{ el: color.el }} onClose={() => setColor(null)} className="pop--color">
          <ColorPicker
            value={color.kind === 'fg' ? st.fg : st.bg}
            noneLabel={color.kind === 'fg' ? 'Обычный цвет текста' : 'Без заливки'}
            onPick={(c) => {
              patch({ [color.kind]: c }, color.kind === 'fg' ? 'Цвет текста' : 'Заливка');
              setColor(null);
            }}
          />
        </Popover>
      )}
    </div>
  );
}

// ─── фото из строки таблицы ──────────────────────────────────────────────────

function LinkedBlock({ img, colName, index, active, onActivate, onOpen }: { img: string; colName: string; index: number; active: boolean; onActivate: () => void; onOpen: () => void }) {
  return (
    <div
      className={'blk blk--linked has-img' + (active ? ' is-active' : '')}
      data-block={index}
      role="group"
      aria-label={`Фото из таблицы, столбец «${colName}»`}
      onPointerDown={onActivate}
    >
      <BlockImage id={img} onOpen={onOpen} />
      <span className="blk-origin">из таблицы · {colName}</span>
    </div>
  );
}

function LinkedToolbar({ colName, onReplace, onRemove }: { colName: string; onReplace: (f: File) => void; onRemove: () => void }) {
  const file = useRef<HTMLInputElement>(null);
  return (
    <div className="blk-tools" role="toolbar" aria-label="Фото из таблицы">
      <button type="button" className="tb" aria-label="Заменить фото в таблице" data-tip="Заменить фото в таблице" onClick={() => file.current?.click()}>
        <ImagePlus {...I} />
      </button>
      <button type="button" className="tb" aria-label="Убрать фото из строки таблицы" data-tip="Убрать фото из строки таблицы" onClick={onRemove}>
        <ImageMinus {...I} />
      </button>
      <span className="blk-tools-note">Это фото из столбца «{colName}» — меняется вместе с таблицей</span>
      <input
        ref={file}
        type="file"
        accept="image/*"
        hidden
        onChange={(e) => {
          const f = e.target.files?.[0];
          e.target.value = '';
          if (f) onReplace(f);
        }}
      />
    </div>
  );
}

// ─── печать этикетки ─────────────────────────────────────────────────────────

function PrintLabel({ sku, title, price }: { sku: string; title: string; price: string }) {
  return createPortal(
    <div className="print-label" aria-hidden>
      <div className="pl-title">{title}</div>
      <Barcode text={sku} className="pl-bars" />
      <div className="pl-sku">{sku}</div>
      {price && <div className="pl-price">{price}</div>}
    </div>,
    document.body,
  );
}

// ─── карточка ────────────────────────────────────────────────────────────────

const MIN_BLOCK_ROWS = 3;

export function ProductCard({ rowId }: { rowId: string }) {
  useStoreVersion();
  const set = useUI((s) => s.set);
  const { sheet, view } = ctx();
  const row: Row | undefined = sheet.rows.get(rowId);
  const [activeBlock, setActiveBlock] = useState<number | null>(null);
  const [more, setMore] = useState<HTMLElement | null>(null);
  const [printing, setPrinting] = useState(false);
  const [uploading, setUploading] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const dlg = useRef<HTMLDivElement>(null);
  // двойной клик по артикулу: второй клик не должен тут же закрыть карточку
  const openedAt = useRef(performance.now());
  const filesInput = useRef<HTMLInputElement>(null);
  const keyHandler = useRef<(e: KeyboardEvent) => void>(() => {});

  // клавиши карточки слушаем на всём документе: фокус может быть где угодно, пока карточка открыта
  useEffect(() => {
    const h = (e: KeyboardEvent) => keyHandler.current(e);
    document.addEventListener('keydown', h);
    return () => document.removeEventListener('keydown', h);
  }, []);

  const close = () => set({ dialog: null });

  useEffect(() => {
    const root = document.getElementById('root');
    root?.setAttribute('inert', '');
    dlg.current?.focus({ preventScroll: true });
    return () => {
      root?.removeAttribute('inert');
      // карточку открывают из таблицы — туда и возвращаем клавиатуру
      requestAnimationFrame(() => gridApi.focus());
    };
  }, []);

  useEffect(() => {
    setActiveBlock(null);
    dlg.current?.querySelector('.card-blocks')?.scrollTo({ top: 0 });
  }, [rowId]);

  useEffect(() => {
    if (!printing) return;
    const done = () => setPrinting(false);
    window.addEventListener('afterprint', done);
    const t = window.setTimeout(() => window.print(), 50);
    return () => {
      window.clearTimeout(t);
      window.removeEventListener('afterprint', done);
    };
  }, [printing]);

  if (!row) return null;

  const phys = sheet.rowOrder.indexOf(rowId);
  const vr = view.rows.indexOf(rowId);
  const keyC = sheet.keyColId ? store.colIndexOf(sheet, sheet.keyColId) : -1;
  const sku = keyC >= 0 ? store.display(sheet, phys, keyC).text : '';
  const title = titleOf(sheet, phys, keyC) || 'Без названия';
  const blocks: CardBlock[] = row.card ?? [];
  // фото из строки таблицы (столбец «Фото» и другие с картинками) — первые клетки сетки.
  // Это та же картинка, что в таблице, не копия: заменили здесь — поменялось в таблице, и наоборот
  const linked = sheet.columns
    .map((col) => ({ col, img: row.cells[col.id]?.img }))
    .filter((x): x is { col: Column; img: string } => !!x.img);
  const L = linked.length;
  const total = Math.max(MIN_BLOCK_ROWS * BLOCKS_PER_ROW, Math.ceil((L + blocks.length) / BLOCKS_PER_ROW) * BLOCKS_PER_ROW);
  const rowsOfBlocks = total / BLOCKS_PER_ROW;
  // собственные блоки карточки идут после фото из таблицы
  const slots: CardBlock[] = Array.from({ length: total - L }, (_, i) => blocks[i] ?? {});
  const photos = L + slots.filter((b) => b.img).length;

  // цена для этикетки: первый столбец в рублях
  const priceC = sheet.columns.findIndex((c, i) => {
    const nf = store.styleOf(sheet, c, sheet.rows.get(rowId)?.cells[c.id]).nf;
    return nf?.k === 'currency' && nf.c === 'RUB' && store.display(sheet, phys, i).text;
  });
  const price = priceC >= 0 ? store.display(sheet, phys, priceC).text : '';

  const go = (delta: number) => {
    const i = vr + delta;
    if (i < 0 || i >= view.rows.length) return;
    set({ dialog: { kind: 'card', rowId: view.rows[i] } });
    const s = useUI.getState().sel;
    setSelection({ ar: i, ac: s.ac, fr: i, fc: s.ac }, true);
  };

  /** Сохранить блоки; хвостовые пустые не храним, но число рядов помним по последнему блоку. */
  const saveSlots = (next: CardBlock[], label: string) => {
    let end = next.length;
    while (end > 0 && !next[end - 1].img && !next[end - 1].text && !next[end - 1].st) end--;
    store.transact(label, () => store.patchRow(sheet, rowId, { card: end ? next.slice(0, end) : undefined }));
  };

  const saveBlock = (i: number, b: CardBlock, label: string) => {
    const next = [...slots];
    next[i] = b;
    saveSlots(next, label);
  };

  /** Несколько фото сразу: первое — в выбранный блок, остальные — в следующие свободные. */
  const addPhotos = async (files: File[], startAt?: number) => {
    const imgs = files.filter(isImageFile);
    if (!imgs.length) {
      toast('Это не изображение. Подойдут JPG, PNG, WebP, HEIC из Safari', { tone: 'error' });
      return;
    }
    const ids: string[] = [];
    try {
      for (let k = 0; k < imgs.length; k++) {
        setUploading(imgs.length > 1 ? `Сжимаем фото ${k + 1} из ${imgs.length}…` : 'Сжимаем фото…');
        ids.push(await importImage(imgs[k], imgs[k].name));
      }
    } catch (e) {
      toast(`Не удалось загрузить фото: ${e instanceof Error ? e.message : e}`, { tone: 'error' });
    } finally {
      setUploading(null);
    }
    if (!ids.length) return;
    let rest = ids;
    if (startAt !== undefined && startAt < L) {
      // перетащили на фото из таблицы — меняем его в самой таблице
      const col = linked[startAt].col;
      store.transact('Фото в таблице', () => store.patchCell(sheet, rowId, col.id, (c) => ({ ...c, img: ids[0] })));
      rest = ids.slice(1);
      if (!rest.length) {
        toast(`Фото в столбце «${col.name}» заменено`);
        return;
      }
    }
    const own = startAt !== undefined && startAt >= L ? startAt - L : undefined;
    const next = [...slots];
    for (let n = 0; n < rest.length; n++) {
      const id = rest[n];
      if (n === 0 && own !== undefined) {
        next[own] = { ...next[own], img: id };
        continue;
      }
      let k = next.findIndex((b, i) => !b.img && i > (own ?? -1));
      if (k < 0) k = next.findIndex((b) => !b.img);
      if (k < 0) {
        for (let add = 0; add < BLOCKS_PER_ROW; add++) next.push({});
        k = next.findIndex((b) => !b.img);
      }
      next[k] = { ...next[k], img: id };
    }
    saveSlots(next, rest.length > 1 ? 'Фото в карточку' : 'Фото в блоке');
    toast(ids.length > 1 ? `Добавлено фото: ${ids.length}` : 'Фото добавлено');
  };

  const addBlockRow = () => {
    const next = [...slots, ...Array.from({ length: BLOCKS_PER_ROW }, () => ({}) as CardBlock)];
    // пустой ряд нужно хранить, иначе он пропадёт: помечаем последний блок пустым оформлением
    next[next.length - 1] = { st: { va: 'bottom' } };
    store.transact('Ряд блоков', () => store.patchRow(sheet, rowId, { card: next }));
  };

  const removeLastBlockRow = () => {
    if (rowsOfBlocks <= MIN_BLOCK_ROWS) return;
    const next = slots.slice(0, Math.max(0, slots.length - BLOCKS_PER_ROW));
    const dropped = slots.slice(next.length);
    if (dropped.some((b) => b.img || b.text) && !confirm('В последнем ряду есть фото или текст. Удалить ряд?')) return;
    store.transact('Удаление ряда блоков', () => store.patchRow(sheet, rowId, { card: next }));
  };

  const deleteRow = () => {
    const label = sku || `строка ${phys + 1}`;
    const nextId = view.rows[vr + 1] ?? view.rows[vr - 1];
    store.deleteRows(sheet, [rowId]);
    toast(`Карточка ${label} удалена`, { action: { label: 'Вернуть', run: () => store.undo() } });
    if (nextId) set({ dialog: { kind: 'card', rowId: nextId } });
    else close();
  };

  const openImage = (id: string) => {
    const list = [...linked.map((x) => x.img), ...slots.map((b) => b.img).filter((x): x is string => !!x)];
    set({ lightbox: { imageId: id, caption: `${sku} · ${title}`, list } });
  };

  keyHandler.current = (e: KeyboardEvent) => {
    if (e.defaultPrevented || document.querySelector('.lb, .pop, .dlg')) return;
    const t = e.target as HTMLElement;
    const typing = t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable;
    if (e.key === 'Escape' && !typing) {
      e.preventDefault();
      close();
    }
    if (typing) return;
    if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
      e.preventDefault();
      go(-1);
    }
    if (e.key === 'ArrowRight' || e.key === 'PageDown') {
      e.preventDefault();
      go(1);
    }
  };

  return createPortal(
    <div className="card-backdrop" onPointerDown={(e) => e.target === e.currentTarget && performance.now() - openedAt.current > 500 && close()}>
      <div ref={dlg} className="card" role="dialog" aria-modal="true" aria-label={`Карточка товара ${sku} — ${title}`} tabIndex={-1}>
        <aside className="card-spine" aria-label="Этикетка">
          <span className="spine-kicker">Артикул</span>
          <div className={'spine-sku' + (sku.length <= 3 ? ' spine-sku--short' : '')} title={sku}>
            {sku || '—'}
          </div>
          {sku && <Barcode text={sku} vertical className="spine-bars spine-bars--v" />}
          {sku && <Barcode text={sku} className="spine-bars spine-bars--h" />}
          <div className="spine-foot">
            <span className="spine-row">стр. {phys + 1}</span>
            <button type="button" className="spine-print" onClick={() => setPrinting(true)} disabled={!sku}>
              <Printer size={15} strokeWidth={1.75} aria-hidden />
              Этикетка
            </button>
          </div>
        </aside>

        <div className="card-main">
          <header className="card-head">
            <div className="card-title">
              <h2>{title}</h2>
              <span className="card-sub">
                {photos ? `фото в карточке: ${photos}` : 'в карточке пока нет фото'}
                {view.filtered && ' · среди отфильтрованных'}
              </span>
            </div>
            <div className="card-nav">
              <button type="button" className="icon-btn" onClick={() => go(-1)} disabled={vr <= 0} aria-label="Предыдущая карточка" title="Предыдущая (←)">
                <ChevronLeft {...I} />
              </button>
              <span className="card-count">
                {vr >= 0 ? (vr + 1).toLocaleString('ru') : '—'} <span>из {view.rows.length.toLocaleString('ru')}</span>
              </span>
              <button type="button" className="icon-btn" onClick={() => go(1)} disabled={vr < 0 || vr >= view.rows.length - 1} aria-label="Следующая карточка" title="Следующая (→)">
                <ChevronRight {...I} />
              </button>
              <button type="button" className="icon-btn" aria-label="Ещё" aria-haspopup="menu" onClick={(e) => setMore(e.currentTarget)}>
                <MoreHorizontal {...I} />
              </button>
              <button type="button" className="icon-btn" onClick={close} aria-label="Закрыть карточку" title="Закрыть (Esc)">
                <X size={18} strokeWidth={1.75} />
              </button>
            </div>
          </header>

          <div className="blk-bar">
            <button type="button" className="btn btn--sm" onClick={() => filesInput.current?.click()} disabled={!!uploading}>
              <ImagePlus size={15} strokeWidth={1.75} aria-hidden />
              {uploading ?? 'Добавить фото'}
            </button>
            <input
              ref={filesInput}
              type="file"
              accept="image/*"
              multiple
              hidden
              onChange={(e) => {
                const files = Array.from(e.target.files ?? []);
                e.target.value = '';
                if (files.length) void addPhotos(files, activeBlock ?? undefined);
              }}
            />
            {activeBlock !== null && activeBlock < L ? (
              <LinkedToolbar
                colName={linked[activeBlock].col.name}
                onReplace={(f) => void addPhotos([f], activeBlock)}
                onRemove={() => {
                  const col = linked[activeBlock].col;
                  store.transact('Убрать фото', () => store.patchCell(sheet, rowId, col.id, (c) => ({ ...c, img: undefined })));
                  setActiveBlock(null);
                }}
              />
            ) : activeBlock !== null ? (
              <BlockToolbar block={slots[activeBlock - L]} onChange={(b, l) => saveBlock(activeBlock - L, b, l)} />
            ) : (
              <span className="blk-bar-hint">Можно выбрать сразу несколько фото или перетащить их на сетку. Выберите блок — появятся выравнивание и цвет.</span>
            )}
          </div>

          <section
            className={'card-blocks' + (dragOver ? ' is-drop' : '')}
            aria-label="Сетка карточки"
            onDragOver={(e) => {
              if (Array.from(e.dataTransfer.types).includes('Files')) {
                e.preventDefault();
                setDragOver(true);
              }
            }}
            onDragLeave={(e) => {
              if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragOver(false);
            }}
            onDrop={(e) => {
              setDragOver(false);
              const files = Array.from(e.dataTransfer.files);
              if (!files.length) return;
              e.preventDefault();
              const blockEl = (e.target as HTMLElement).closest<HTMLElement>('[data-block]');
              void addPhotos(files, blockEl ? Number(blockEl.dataset.block) : undefined);
            }}
          >
            <div className="blk-grid" style={{ '--rows': rowsOfBlocks } as React.CSSProperties}>
              {linked.map((x, i) => (
                <LinkedBlock key={'t' + x.col.id} img={x.img} colName={x.col.name} index={i} active={activeBlock === i} onActivate={() => setActiveBlock(i)} onOpen={() => openImage(x.img)} />
              ))}
              {slots.map((b, i) => (
                <Block
                  key={i}
                  block={b}
                  index={i + L}
                  active={activeBlock === i + L}
                  onActivate={() => setActiveBlock(i + L)}
                  onChange={(nb, l) => saveBlock(i, nb, l)}
                  onOpenImage={openImage}
                />
              ))}
            </div>
            <div className="blk-actions">
              <button type="button" className="btn btn--ghost btn--sm" onClick={addBlockRow}>
                <Plus size={14} strokeWidth={2} aria-hidden /> Ряд блоков
              </button>
              {rowsOfBlocks > MIN_BLOCK_ROWS && (
                <button type="button" className="btn btn--ghost btn--sm" onClick={removeLastBlockRow}>
                  Убрать последний ряд
                </button>
              )}
            </div>
          </section>
        </div>
        {more && (
          <Menu
            anchor={{ el: more }}
            placement="bottom-end"
            onClose={() => setMore(null)}
            items={[
              {
                label: 'Показать в таблице',
                onSelect: () => {
                  if (vr >= 0) setSelection({ ar: vr, ac: 0, fr: vr, fc: 0 });
                  close();
                },
              },
              { label: 'Печать этикетки', icon: <Printer size={15} strokeWidth={1.75} />, disabled: !sku, onSelect: () => setPrinting(true) },
              'sep',
              { label: 'Удалить карточку (строку)', icon: <Trash2 size={15} strokeWidth={1.75} />, danger: true, onSelect: deleteRow },
            ]}
          />
        )}
      </div>
      {printing && <PrintLabel sku={sku} title={title} price={price} />}
    </div>,
    document.body,
  );
}
