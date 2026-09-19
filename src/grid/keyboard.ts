import {
  cellAt,
  clearSelection,
  ctx,
  deleteSelectedRows,
  fillDown,
  fillRight,
  insertRows,
  jump,
  moveSel,
  openCard,
  redo,
  selectAll,
  setSelection,
  startEdit,
  toggleStyle,
  undo,
} from '../app/actions';
import { gridApi } from '../app/gridApi';
import { useUI } from '../ui/state';

const mod = (e: KeyboardEvent | React.KeyboardEvent) => e.ctrlKey || e.metaKey;

/** Клавиши сетки вне режима редактирования. Возвращает true, если клавиша обработана. */
export function handleGridKey(e: React.KeyboardEvent): boolean {
  const ui = useUI.getState();
  const s = ui.sel;
  const k = e.key;
  const shift = e.shiftKey;

  if (mod(e)) {
    const lower = k.toLowerCase();
    // русская раскладка: Ctrl+Я = Ctrl+Z и т.д. — сравниваем по физической клавише
    const code = e.code;
    if (code === 'KeyZ' && !shift) return run(undo);
    if ((code === 'KeyZ' && shift) || code === 'KeyY') return run(redo);
    if (code === 'KeyB') return run(() => toggleStyle('b'));
    if (code === 'KeyI') return run(() => toggleStyle('i'));
    if (code === 'KeyU') return run(() => toggleStyle('u'));
    if (code === 'Digit5') return run(() => toggleStyle('s'));
    if (code === 'KeyA') return run(selectAll);
    if (code === 'KeyD') return run(fillDown);
    if (code === 'KeyR') return run(fillRight);
    if (code === 'KeyF') return run(() => ui.set({ searchOpen: true }));
    if (code === 'KeyK') return run(() => ui.set({ dialog: { kind: 'link', r: s.ar, c: s.ac } }));
    if (code === 'Space') return run(() => selectColumn());
    if (lower === 'arrowdown') return run(() => jump(1, 0, shift));
    if (lower === 'arrowup') return run(() => jump(-1, 0, shift));
    if (lower === 'arrowright') return run(() => jump(0, 1, shift));
    if (lower === 'arrowleft') return run(() => jump(0, -1, shift));
    if (k === 'Home') return run(() => setSelection(shift ? { ...s, fr: 0, fc: 0 } : { ar: 0, ac: 0, fr: 0, fc: 0 }));
    if (k === 'End') {
      const { view } = ctx();
      const r = view.rows.length - 1;
      const c = view.cols.length - 1;
      return run(() => setSelection(shift ? { ...s, fr: r, fc: c } : { ar: r, ac: c, fr: r, fc: c }));
    }
    if (k === 'Enter') return run(() => openCard(s.ar));
    if (k === '-' || code === 'Minus') return run(deleteSelectedRows);
    if (k === '+' || k === '=' || code === 'Equal') return run(() => insertRows('above'));
    return false;
  }

  switch (k) {
    case 'ArrowDown':
      return run(() => moveSel(1, 0, shift));
    case 'ArrowUp':
      return run(() => moveSel(-1, 0, shift));
    case 'ArrowRight':
      return run(() => moveSel(0, 1, shift));
    case 'ArrowLeft':
      return run(() => moveSel(0, -1, shift));
    case 'Tab':
      return run(() => moveSel(0, shift ? -1 : 1));
    case 'Enter':
      if (e.altKey) return false;
      return run(() => moveSel(shift ? -1 : 1, 0));
    case 'PageDown':
      return run(() => moveSel(gridApi.pageRows(), 0, shift));
    case 'PageUp':
      return run(() => moveSel(-gridApi.pageRows(), 0, shift));
    case 'Home':
      return run(() => setSelection(shift ? { ...s, fc: 0 } : { ar: s.ar, ac: 0, fr: s.ar, fc: 0 }));
    case 'End': {
      const { view } = ctx();
      const c = view.cols.length - 1;
      return run(() => setSelection(shift ? { ...s, fc: c } : { ar: s.ar, ac: c, fr: s.ar, fc: c }));
    }
    case 'F2':
      return run(() => startEdit('edit'));
    case 'Delete':
    case 'Backspace':
      if (k === 'Backspace') {
        // Backspace в Excel очищает ячейку и начинает ввод
        const x = cellAt(s.ar, s.ac);
        if (x && s.ar === s.fr && s.ac === s.fc) return run(() => startEdit('enter', ''));
      }
      return run(() => clearSelection('contents'));
    case 'Escape':
      if (ui.copyRange) return run(() => ui.set({ copyRange: null }));
      return false;
    case ' ':
      if (shift) return run(selectRow);
      break;
  }
  // печатный символ — начать ввод поверх
  if (k.length === 1 && !e.altKey) {
    startEdit('enter', k);
    e.preventDefault();
    return true;
  }
  return false;

  function run(fn: () => void) {
    fn();
    e.preventDefault();
    return true;
  }
}

function selectColumn() {
  const s = useUI.getState().sel;
  const { view } = ctx();
  setSelection({ ar: 0, ac: s.ac, fr: view.rows.length - 1, fc: s.fc }, false);
}

function selectRow() {
  const s = useUI.getState().sel;
  const { view } = ctx();
  setSelection({ ar: s.ar, ac: 0, fr: s.fr, fc: view.cols.length - 1 }, false);
}

export function isFormulaPointMode(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false;
  const before = text.slice(0, caret).trimEnd();
  return /[=(;,+\-*/^&<>:]$/.test(before);
}

