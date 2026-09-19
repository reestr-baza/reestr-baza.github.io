import {
  ArrowDownWideNarrow,
  ArrowUpNarrowWide,
  ClipboardPaste,
  Copy,
  Eraser,
  Eye,
  EyeOff,
  ImageMinus,
  ImagePlus,
  KeyRound,
  Link2,
  PanelLeftClose,
  PanelsTopLeft,
  PenLine,
  Scissors,
  SquareArrowOutUpRight,
  SquareSigma,
  StickyNote,
  Trash2,
} from 'lucide-react';
import { useRef } from 'react';
import type { NumFmt } from '../model/types';
import { Menu, type MenuItem } from '../ui/Menu';
import { selRect, useUI } from '../ui/state';
import {
  applyColumnFormula,
  cellAt,
  clearSelection,
  ctx,
  deleteSelectedColumns,
  deleteSelectedRows,
  hideSelectedColumns,
  hideSelectedRows,
  insertColumns,
  setNumFmt,
  insertImages,
  insertRows,
  openCard,
  removeImage,
  sortByColumn,
  toast,
  unhideColumnsAround,
  unhideRowsAround,
} from './actions';
import { gridApi } from './gridApi';
import { store } from './instance';

const I = { size: 15, strokeWidth: 1.75 };

/** Быстрые форматы в меню заголовка: ₽, ¥, $, %, число, общий. */
function colFormats(nf: NumFmt | undefined): MenuItem[] {
  const is = (k: string, c?: string) => nf?.k === k && (!c || (nf.k === 'currency' && nf.c === c));
  return [
    { label: 'Рубли  ₽', checked: is('currency', 'RUB'), onSelect: () => setNumFmt({ k: 'currency', c: 'RUB', d: 0 }) },
    { label: 'Юани  ¥', checked: is('currency', 'CNY'), onSelect: () => setNumFmt({ k: 'currency', c: 'CNY', d: 0 }) },
    { label: 'Доллары  $', checked: is('currency', 'USD'), onSelect: () => setNumFmt({ k: 'currency', c: 'USD', d: 2 }) },
    { label: 'Процент  %', checked: is('percent'), onSelect: () => setNumFmt({ k: 'percent', d: 0 }) },
    { label: 'Число', checked: is('number'), onSelect: () => setNumFmt({ k: 'number', d: 2 }) },
    { label: 'Дата', checked: is('date'), onSelect: () => setNumFmt({ k: 'date' }) },
    { label: 'Общий', checked: !nf || nf.k === 'general', onSelect: () => setNumFmt(undefined) },
  ];
}

function execClipboard(cmd: 'copy' | 'cut') {
  gridApi.focus();
  if (!document.execCommand(cmd)) toast(`Используйте Ctrl+${cmd === 'copy' ? 'C' : 'X'}`);
}

async function pasteFromMenu() {
  gridApi.focus();
  // браузеры не дают вставлять из меню без разрешения — подсказываем сочетание
  toast('Чтобы вставить, нажмите Ctrl+V');
}

export function ContextMenus() {
  const menu = useUI((s) => s.menu);
  const set = useUI((s) => s.set);
  const file = useRef<HTMLInputElement>(null);
  const close = () => set({ menu: null });

  const picker = (
    <input
      ref={file}
      type="file"
      accept="image/*"
      multiple
      hidden
      onChange={(e) => {
        const files = Array.from(e.target.files ?? []);
        e.target.value = '';
        const s = useUI.getState().sel;
        if (files.length) void insertImages(files, s.ar, s.ac);
      }}
    />
  );
  if (!menu) return picker;

  const { sheet, view } = ctx();
  const sel = useUI.getState().sel;
  const { r1, r2, c1, c2 } = selRect(sel);
  const nRows = r2 - r1 + 1;
  const nCols = c2 - c1 + 1;
  const rowsTxt = nRows > 1 ? ` (${nRows})` : '';
  const colsTxt = nCols > 1 ? ` (${nCols})` : '';
  let items: MenuItem[] = [];

  if (menu.kind === 'cell') {
    const x = cellAt(sel.ar, sel.ac);
    const hasImg = !!x?.row?.cells[x.colId]?.img;
    items = [
      { label: 'Вырезать', icon: <Scissors {...I} />, hint: 'Ctrl+X', onSelect: () => execClipboard('cut') },
      { label: 'Копировать', icon: <Copy {...I} />, hint: 'Ctrl+C', onSelect: () => execClipboard('copy') },
      { label: 'Вставить', icon: <ClipboardPaste {...I} />, hint: 'Ctrl+V', onSelect: pasteFromMenu },
      'sep',
      { label: 'Открыть карточку', icon: <SquareArrowOutUpRight {...I} />, hint: 'Ctrl+Enter', onSelect: () => openCard(sel.ar) },
      { label: hasImg ? 'Заменить фото…' : 'Вставить фото…', icon: <ImagePlus {...I} />, onSelect: () => file.current?.click() },
      ...(hasImg ? [{ label: 'Убрать фото', icon: <ImageMinus {...I} />, onSelect: removeImage }] : []),
      { label: 'Ссылка…', icon: <Link2 {...I} />, hint: 'Ctrl+K', onSelect: () => set({ dialog: { kind: 'link', r: sel.ar, c: sel.ac } }) },
      { label: x?.row?.cells[x.colId]?.note ? 'Изменить примечание…' : 'Примечание…', icon: <StickyNote {...I} />, onSelect: () => set({ dialog: { kind: 'note', r: sel.ar, c: sel.ac } }) },
      'sep',
      { label: `Вставить строки выше${rowsTxt}`, onSelect: () => insertRows('above') },
      { label: `Вставить строки ниже${rowsTxt}`, onSelect: () => insertRows('below') },
      { label: `Удалить строки${rowsTxt}`, icon: <Trash2 {...I} />, danger: true, onSelect: deleteSelectedRows },
      { label: `Скрыть строки${rowsTxt}`, icon: <EyeOff {...I} />, onSelect: hideSelectedRows },
      'sep',
      { label: 'Очистить содержимое', icon: <Eraser {...I} />, hint: 'Delete', onSelect: () => clearSelection('contents') },
      { label: 'Очистить формат', onSelect: () => clearSelection('formats') },
    ];
  } else if (menu.kind === 'col') {
    const col = sheet.columns[view.cols[menu.c]];
    const isKey = sheet.keyColId === col.id;
    const frozenHere = sheet.frozen === menu.c + 1;
    items = [
      { label: 'Переименовать…', icon: <PenLine {...I} />, onSelect: () => set({ dialog: { kind: 'rename-col', c: menu.c } }) },
      { label: 'Сортировать по возрастанию', icon: <ArrowUpNarrowWide {...I} />, onSelect: () => sortByColumn(menu.c, 'asc') },
      { label: 'Сортировать по убыванию', icon: <ArrowDownWideNarrow {...I} />, onSelect: () => sortByColumn(menu.c, 'desc') },
      'sep',
      { label: `Вставить столбцы слева${colsTxt}`, onSelect: () => insertColumns('left') },
      { label: `Вставить столбцы справа${colsTxt}`, onSelect: () => insertColumns('right') },
      { label: `Удалить столбцы${colsTxt}`, icon: <Trash2 {...I} />, danger: true, onSelect: deleteSelectedColumns },
      { label: `Скрыть столбцы${colsTxt}`, icon: <EyeOff {...I} />, onSelect: hideSelectedColumns },
      { label: 'Показать скрытые рядом', icon: <Eye {...I} />, onSelect: unhideColumnsAround },
      'sep',
      { heading: 'Формат столбца' },
      ...colFormats(col.st?.nf),
      'sep',
      {
        label: isKey ? 'Ключевой столбец карточки' : 'Открывать карточку по этому столбцу',
        icon: <KeyRound {...I} />,
        checked: isKey,
        onSelect: () => store.transact('Ключевой столбец', () => store.setSheetMeta(sheet, { keyColId: isKey ? undefined : col.id })),
      },
      {
        label: frozenHere ? 'Открепить столбцы' : 'Закрепить столбцы до этого',
        icon: frozenHere ? <PanelLeftClose {...I} /> : <PanelsTopLeft {...I} />,
        onSelect: () => store.transact('Закрепление', () => store.setSheetMeta(sheet, { frozen: frozenHere ? 0 : menu.c + 1 })),
      },
      ...(col.formula
        ? [{ label: 'Убрать формулу столбца', icon: <SquareSigma {...I} />, onSelect: () => applyColumnFormula(col.id, undefined) }]
        : []),
    ];
  } else if (menu.kind === 'row') {
    items = [
      { label: 'Открыть карточку', icon: <SquareArrowOutUpRight {...I} />, onSelect: () => openCard(menu.r) },
      'sep',
      { label: `Вставить строки выше${rowsTxt}`, onSelect: () => insertRows('above') },
      { label: `Вставить строки ниже${rowsTxt}`, onSelect: () => insertRows('below') },
      { label: `Удалить строки${rowsTxt}`, icon: <Trash2 {...I} />, danger: true, onSelect: deleteSelectedRows },
      'sep',
      { label: `Скрыть строки${rowsTxt}`, icon: <EyeOff {...I} />, onSelect: hideSelectedRows },
      { label: 'Показать скрытые рядом', icon: <Eye {...I} />, onSelect: unhideRowsAround },
      { label: 'Высота по умолчанию', onSelect: () => store.setRowsHeight(sheet, view.rows.slice(r1, r2 + 1), undefined) },
    ];
  } else if (menu.kind === 'sheet') {
    const target = store.sheets.get(menu.sheetId);
    const many = store.meta.sheetIds.length > 1;
    items = [
      {
        label: 'Удалить лист',
        icon: <Trash2 {...I} />,
        danger: true,
        disabled: !many,
        onSelect: () => {
          if (!target) return;
          store.removeSheet(target.id);
          set({ sel: { ar: 0, ac: 0, fr: 0, fc: 0 } });
          toast(`Лист «${target.name}» удалён`, { action: { label: 'Вернуть', run: () => store.undo() } });
        },
      },
    ];
  }

  return (
    <>
      {picker}
      <Menu anchor={{ x: menu.x, y: menu.y }} items={items} onClose={close} label="Контекстное меню" />
    </>
  );
}
