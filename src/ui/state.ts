import { create } from 'zustand';
import type { RowId } from '../model/types';

/** Выделение в координатах представления: строки — видимые строки, столбцы — видимые столбцы. */
export interface Selection {
  /** Активная ячейка */
  ar: number;
  ac: number;
  /** Противоположный угол диапазона */
  fr: number;
  fc: number;
}

export interface EditState {
  r: number;
  c: number;
  text: string;
  /** enter — начали печатать поверх (стрелки завершают ввод); edit — F2/двойной клик */
  mode: 'enter' | 'edit';
  source: 'cell' | 'bar';
  /** Сдвиг курсора после вставки ссылки */
  caret?: number;
}

export type MenuState =
  | { kind: 'cell'; x: number; y: number }
  | { kind: 'col'; x: number; y: number; c: number }
  | { kind: 'row'; x: number; y: number; r: number }
  | { kind: 'sheet'; x: number; y: number; sheetId: string }
  | null;

export type DialogState =
  | { kind: 'card'; rowId: RowId }
  | { kind: 'link'; r: number; c: number }
  | { kind: 'import' }
  | { kind: 'backup' }
  | { kind: 'settings' }
  | { kind: 'shortcuts' }
  | { kind: 'rename-col'; c: number }
  | { kind: 'note'; r: number; c: number }
  | null;

export interface Toast {
  id: number;
  text: string;
  tone?: 'error' | 'plain';
  action?: { label: string; run: () => void };
}

interface UIState {
  sel: Selection;
  edit: EditState | null;
  menu: MenuState;
  filterMenu: { c: number; rect: { left: number; top: number; bottom: number; right: number } } | null;
  dialog: DialogState;
  /** Просмотр фото — поверх карточки или таблицы */
  lightbox: { imageId: string; caption?: string; list?: string[] } | null;
  toasts: Toast[];
  searchOpen: boolean;
  /** Внутренний буфер обмена: помечает, что копировали мы, и хранит формулы/стили/фото */
  clipboardId: string | null;
  /** «Бегущие муравьи» вокруг скопированного диапазона */
  copyRange: { r1: number; c1: number; r2: number; c2: number; cut: boolean } | null;
  set: (patch: Partial<UIState>) => void;
  setSel: (sel: Selection) => void;
  toast: (t: Omit<Toast, 'id'>) => void;
  dismissToast: (id: number) => void;
}

let toastSeq = 1;

export const useUI = create<UIState>((set) => ({
  sel: { ar: 0, ac: 0, fr: 0, fc: 0 },
  edit: null,
  menu: null,
  filterMenu: null,
  dialog: null,
  lightbox: null,
  toasts: [],
  searchOpen: false,
  clipboardId: null,
  copyRange: null,
  set: (patch) => set(patch),
  setSel: (sel) => set({ sel }),
  toast: (t) => {
    const id = toastSeq++;
    set((s) => ({ toasts: [...s.toasts.slice(-2), { ...t, id }] }));
    if (!t.action && t.tone !== 'error') {
      window.setTimeout(() => set((s) => ({ toasts: s.toasts.filter((x) => x.id !== id) })), 3800);
    } else {
      window.setTimeout(() => set((s) => ({ toasts: s.toasts.filter((x) => x.id !== id) })), 9000);
    }
  },
  dismissToast: (id) => set((s) => ({ toasts: s.toasts.filter((x) => x.id !== id) })),
}));

export function selRect(s: Selection) {
  return {
    r1: Math.min(s.ar, s.fr),
    r2: Math.max(s.ar, s.fr),
    c1: Math.min(s.ac, s.fc),
    c2: Math.max(s.ac, s.fc),
  };
}
