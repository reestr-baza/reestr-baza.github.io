/** Мост между действиями и смонтированной сеткой: прокрутка, фокус, координаты. */
export const gridApi = {
  scrollToCell: (_vr: number, _vc: number) => {},
  focus: () => {},
  /** Прямоугольник ячейки на экране — для меню и подсказок. */
  cellRect: (_vr: number, _vc: number): DOMRect | null => null,
  /** Видимая высота в строках — для PageUp/PageDown. */
  pageRows: () => 20,
};
