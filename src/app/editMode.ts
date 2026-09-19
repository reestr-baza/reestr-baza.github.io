import { useUI } from '../ui/state';
import { store } from './instance';

const HINT = 'Сейчас режим просмотра — менять данные нельзя';

/**
 * Просмотр и редактирование.
 * База открывается в просмотре: искать, фильтровать, сортировать и открывать карточки можно,
 * а менять данные — только после кнопки «Редактировать».
 */
export function setEditing(on: boolean) {
  store.readOnly = !on;
  const ui = useUI.getState();
  // подсказка «сейчас режим просмотра» больше не актуальна
  ui.set(on ? { editing: true, toasts: ui.toasts.filter((t) => t.text !== HINT) } : { editing: false, edit: null });
}

export const isEditing = () => useUI.getState().editing;

/** Можно ли менять данные. Если нет — подсказываем, как включить редактирование. */
export function requireEdit(): boolean {
  if (isEditing()) return true;
  const ui = useUI.getState();
  // не заваливаем подсказками, если человек продолжает печатать
  if (!ui.toasts.some((t) => t.text === HINT)) ui.toast({ text: HINT, action: { label: 'Редактировать', run: () => setEditing(true) } });
  return false;
}
