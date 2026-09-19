import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { appendRow, cellAt, commitEdit, ctx, fillTo, insertImages, openCard, setSelection, startEdit } from '../app/actions';
import { handleCopy, handlePaste } from '../app/clipboard';
import { gridApi } from '../app/gridApi';
import { store, useStoreVersion } from '../app/instance';
import { colToLetters } from '../formula/a1';
import { listRefs } from '../formula/refs';
import { isImageFile } from '../storage/images';
import { selRect, useUI } from '../ui/state';
import { CellEditor } from './Editor';
import { buildGeometry, HD_H, indexAt, RH_W, type Geometry } from './geometry';
import { handleGridKey, isFormulaPointMode } from './keyboard';
import { GridRow } from './Row';

const ADD_ROW_H = 40;
const ADD_COL_W = 44;
const OVERSCAN_ROWS = 6;

type Zone =
  | { z: 'cell'; vr: number; vc: number }
  | { z: 'col'; vc: number }
  | { z: 'row'; vr: number }
  | { z: 'corner' }
  | { z: 'none' };

type Drag =
  | { kind: 'cells'; pointerId: number }
  | { kind: 'cols'; pointerId: number }
  | { kind: 'rows'; pointerId: number }
  | { kind: 'fill'; pointerId: number; axis: 'row' | 'col' | null; to: number }
  | { kind: 'colResize'; pointerId: number; c: number; startX: number; startW: number; w: number }
  | { kind: 'rowResize'; pointerId: number; vr: number; startY: number; startH: number; h: number }
  | { kind: 'point'; pointerId: number; anchor: { vr: number; vc: number }; start: number; end: number }
  | null;

const REF_COLORS = ['#2f5fd0', '#c2410c', '#15803d', '#a21caf', '#0e7490', '#b45309'];

export function Grid() {
  useStoreVersion();
  const sheet = store.activeSheet;
  const view = store.view(sheet);
  const sel = useUI((s) => s.sel);
  const edit = useUI((s) => s.edit);
  const copyRange = useUI((s) => s.copyRange);
  const filterMenu = useUI((s) => s.filterMenu);

  const [vp, setVp] = useState({ top: 0, left: 0, w: 1200, h: 800 });
  const widthStep = Math.round(vp.w / 40);
  const geo = useMemo(
    () => buildGeometry(store, sheet, view, Math.max(120, widthStep * 40 * 0.55 - RH_W)),
    // layoutVersion покрывает размеры, порядок и видимость; view — фильтры
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [store.layoutVersion, view, sheet, sheet.density, sheet.frozen, sheet.columns, widthStep],
  );
  const geoRef = useRef(geo);
  geoRef.current = geo;

  const scroller = useRef<HTMLDivElement>(null);
  const raf = useRef(0);
  const [drag, setDragState] = useState<Drag>(null);
  const dragRef = useRef<Drag>(null);
  const setDrag = (d: Drag) => {
    dragRef.current = d;
    setDragState(d);
  };
  const [dropTarget, setDropTarget] = useState<{ vr: number; vc: number } | null>(null);

  const leftW = RH_W + geo.frozenW;
  const mainW = geo.totalW - geo.frozenW + ADD_COL_W;

  // ─── окно просмотра ────────────────────────────────────────────────────────
  const readViewport = useCallback(() => {
    const el = scroller.current;
    if (!el) return;
    setVp((p) =>
      p.top === el.scrollTop && p.left === el.scrollLeft && p.w === el.clientWidth && p.h === el.clientHeight
        ? p
        : { top: el.scrollTop, left: el.scrollLeft, w: el.clientWidth, h: el.clientHeight },
    );
  }, []);

  useLayoutEffect(() => {
    const el = scroller.current!;
    readViewport();
    const ro = new ResizeObserver(readViewport);
    ro.observe(el);
    return () => ro.disconnect();
  }, [readViewport]);

  const onScroll = () => {
    if (raf.current) return;
    raf.current = requestAnimationFrame(() => {
      raf.current = 0;
      readViewport();
    });
  };

  // при смене листа — наверх
  useEffect(() => {
    scroller.current?.scrollTo({ top: 0, left: 0 });
  }, [sheet.id]);

  // ─── API для действий ──────────────────────────────────────────────────────
  useEffect(() => {
    gridApi.scrollToCell = (vr, vc) => {
      const el = scroller.current;
      const g = geoRef.current;
      if (!el || vr < 0 || vr >= g.phys.length) return;
      const bodyH = el.clientHeight - HD_H;
      const y0 = g.rowY[vr];
      const y1 = g.rowY[vr + 1];
      let top = el.scrollTop;
      if (y0 < top) top = y0;
      else if (y1 > top + bodyH) top = y1 - bodyH;
      let left = el.scrollLeft;
      if (vc >= g.frozen && vc < g.cols.length) {
        const mainVis = el.clientWidth - RH_W - g.frozenW;
        const x0 = g.colX[vc] - g.frozenW;
        const x1 = g.colX[vc + 1] - g.frozenW;
        if (x0 < left) left = x0;
        else if (x1 > left + mainVis) left = Math.min(x0, x1 - mainVis);
      }
      if (top !== el.scrollTop || left !== el.scrollLeft) el.scrollTo({ top, left });
    };
    gridApi.focus = () => scroller.current?.focus({ preventScroll: true });
    gridApi.cellRect = (vr, vc) => {
      const el = scroller.current;
      const g = geoRef.current;
      if (!el || vr < 0 || vr >= g.phys.length || vc < 0 || vc >= g.cols.length) return null;
      const r = el.getBoundingClientRect();
      const x =
        vc < g.frozen ? r.left + RH_W + g.colX[vc] : r.left + RH_W + g.frozenW + (g.colX[vc] - g.frozenW - el.scrollLeft);
      const y = r.top + HD_H + g.rowY[vr] - el.scrollTop;
      return new DOMRect(x, y, g.colX[vc + 1] - g.colX[vc], g.rowY[vr + 1] - g.rowY[vr]);
    };
    gridApi.pageRows = () => {
      const el = scroller.current;
      const g = geoRef.current;
      if (!el || !g.phys.length) return 20;
      const avg = g.totalH / g.phys.length;
      return Math.max(1, Math.floor((el.clientHeight - HD_H) / avg) - 1);
    };
  }, []);

  // выделение для отмены/повтора
  useEffect(() => {
    store.selectionProvider = {
      get: () => ({ sheetId: store.activeSheet.id, sel: useUI.getState().sel }),
      set: (v) => {
        const x = v as { sheetId: string; sel: typeof sel };
        if (x?.sheetId === store.activeSheet.id) useUI.getState().setSel(x.sel);
      },
    };
  }, []);

  // ─── буфер обмена ──────────────────────────────────────────────────────────
  useEffect(() => {
    const owns = () => document.activeElement === scroller.current && !useUI.getState().edit;
    const onCopy = (e: ClipboardEvent) => owns() && handleCopy(e, false);
    const onCut = (e: ClipboardEvent) => owns() && handleCopy(e, true);
    const onPaste = (e: ClipboardEvent) => owns() && handlePaste(e);
    document.addEventListener('copy', onCopy);
    document.addEventListener('cut', onCut);
    document.addEventListener('paste', onPaste);
    return () => {
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCut);
      document.removeEventListener('paste', onPaste);
    };
  }, []);

  // ─── попадание указателя ───────────────────────────────────────────────────
  const hit = useCallback((clientX: number, clientY: number, clamp = false): Zone => {
    const el = scroller.current;
    const g = geoRef.current;
    if (!el) return { z: 'none' };
    const r = el.getBoundingClientRect();
    const px = clientX - r.left;
    const py = clientY - r.top;
    const inHead = py < HD_H;
    const lw = RH_W + g.frozenW;
    let dataX: number;
    if (px < RH_W && !clamp) {
      if (inHead) return { z: 'corner' };
      const y = py - HD_H + el.scrollTop;
      if (y >= g.totalH) return { z: 'none' };
      return { z: 'row', vr: indexAt(g.rowY, y) };
    }
    if (px < lw) dataX = Math.max(0, px - RH_W);
    else dataX = px - lw + el.scrollLeft + g.frozenW;
    const beyondCols = dataX >= g.totalW;
    const vc = Math.max(0, Math.min(g.cols.length - 1, indexAt(g.colX, dataX)));
    if (inHead && !clamp) return beyondCols ? { z: 'none' } : { z: 'col', vc };
    const y = Math.max(0, py - HD_H + el.scrollTop);
    if (!clamp && (y >= g.totalH || beyondCols)) return { z: 'none' };
    const vr = Math.max(0, Math.min(g.phys.length - 1, indexAt(g.rowY, y)));
    return { z: 'cell', vr, vc };
  }, []);

  // ─── автопрокрутка при перетаскивании ──────────────────────────────────────
  const lastPointer = useRef({ x: 0, y: 0 });
  const autoScroll = useRef(0);
  const stepAutoScroll = useCallback(() => {
    const el = scroller.current;
    if (!el || !dragRef.current) {
      autoScroll.current = 0;
      return;
    }
    const r = el.getBoundingClientRect();
    const { x, y } = lastPointer.current;
    const g = geoRef.current;
    let dy = 0;
    let dx = 0;
    if (y < r.top + HD_H + 8) dy = -Math.min(40, r.top + HD_H + 8 - y);
    else if (y > r.bottom - 12) dy = Math.min(40, y - (r.bottom - 12));
    if (x < r.left + RH_W + g.frozenW + 8) dx = -Math.min(40, r.left + RH_W + g.frozenW + 8 - x);
    else if (x > r.right - 12) dx = Math.min(40, x - (r.right - 12));
    if (dx || dy) {
      el.scrollBy(dx, dy);
      handleDragMove(x, y);
      autoScroll.current = requestAnimationFrame(stepAutoScroll);
    } else autoScroll.current = 0;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ─── перетаскивание ────────────────────────────────────────────────────────
  const handleDragMove = (clientX: number, clientY: number) => {
    const d = dragRef.current;
    if (!d) return;
    const s = useUI.getState().sel;
    const g = geoRef.current;
    if (d.kind === 'cells' || d.kind === 'point') {
      const h = hit(clientX, clientY, true);
      if (h.z !== 'cell') return;
      if (d.kind === 'cells') {
        if (h.vr !== s.fr || h.vc !== s.fc) useUI.getState().setSel({ ...s, fr: h.vr, fc: h.vc });
      } else updatePointRef(d, h.vr, h.vc);
    } else if (d.kind === 'cols') {
      const h = hit(clientX, Math.max(clientY, 0), true);
      if (h.z === 'cell' && h.vc !== s.fc) useUI.getState().setSel({ ...s, fc: h.vc });
    } else if (d.kind === 'rows') {
      const el = scroller.current!;
      const r = el.getBoundingClientRect();
      const y = Math.max(0, clientY - r.top - HD_H + el.scrollTop);
      const vr = Math.min(g.phys.length - 1, indexAt(g.rowY, y));
      if (vr !== s.fr) useUI.getState().setSel({ ...s, fr: vr });
    } else if (d.kind === 'fill') {
      const h = hit(clientX, clientY, true);
      if (h.z !== 'cell') return;
      const { r1, r2, c1, c2 } = selRect(s);
      const outRow = h.vr > r2 ? h.vr - r2 : h.vr < r1 ? r1 - h.vr : 0;
      const outCol = h.vc > c2 ? h.vc - c2 : h.vc < c1 ? c1 - h.vc : 0;
      let next: Drag;
      if (!outRow && !outCol) next = { ...d, axis: null, to: 0 };
      else if (outRow >= outCol) next = { ...d, axis: 'row', to: h.vr };
      else next = { ...d, axis: 'col', to: h.vc };
      if (next.axis !== d.axis || next.to !== d.to) setDrag(next);
    } else if (d.kind === 'colResize') {
      const w = Math.max(36, Math.min(900, d.startW + clientX - d.startX));
      if (w !== d.w) setDrag({ ...d, w });
    } else if (d.kind === 'rowResize') {
      const h = Math.max(20, Math.min(600, d.startH + clientY - d.startY));
      if (h !== d.h) setDrag({ ...d, h });
    }
  };

  const updatePointRef = (d: Extract<Drag, { kind: 'point' }>, vr: number, vc: number) => {
    const ed = useUI.getState().edit;
    if (!ed) return;
    const g = geoRef.current;
    const a = d.anchor;
    const refOf = (r: number, c: number) => `${colToLetters(g.cols[c])}${g.phys[r] + 1}`;
    const ref = a.vr === vr && a.vc === vc ? refOf(vr, vc) : `${refOf(Math.min(a.vr, vr), Math.min(a.vc, vc))}:${refOf(Math.max(a.vr, vr), Math.max(a.vc, vc))}`;
    const text = ed.text.slice(0, d.start) + ref + ed.text.slice(d.end);
    const end = d.start + ref.length;
    useUI.getState().set({ edit: { ...ed, text, caret: end } });
    const nd: Drag = { ...d, end };
    dragRef.current = nd;
  };

  const endDrag = () => {
    const d = dragRef.current;
    if (!d) return;
    const { sheet } = ctx();
    if (d.kind === 'fill' && d.axis) fillTo(d.axis, d.to);
    if (d.kind === 'colResize') {
      const col = sheet.columns[geoRef.current.cols[d.c]];
      if (col && col.w !== d.w) {
        // ширину меняем сразу всем выделенным столбцам, если тянут один из них
        const s = selRect(useUI.getState().sel);
        const all = selectionIsWholeCols() && d.c >= s.c1 && d.c <= s.c2;
        const ids = all ? geoRef.current.cols.slice(s.c1, s.c2 + 1).map((c) => sheet.columns[c].id) : [col.id];
        store.transact('Ширина столбца', () => ids.forEach((id) => store.updateColumn(sheet, id, { w: d.w })));
      }
    }
    if (d.kind === 'rowResize') {
      const rowId = geoRef.current.view.rows[d.vr];
      const s = selRect(useUI.getState().sel);
      const all = selectionIsWholeRows() && d.vr >= s.r1 && d.vr <= s.r2;
      const ids = all ? geoRef.current.view.rows.slice(s.r1, s.r2 + 1) : [rowId];
      store.setRowsHeight(sheet, ids, d.h);
    }
    if (d.kind === 'point') pointInput.current?.focus();
    setDrag(null);
    cancelAnimationFrame(autoScroll.current);
    autoScroll.current = 0;
  };

  useEffect(() => {
    if (!drag) return;
    const move = (e: PointerEvent) => {
      lastPointer.current = { x: e.clientX, y: e.clientY };
      handleDragMove(e.clientX, e.clientY);
      if (!autoScroll.current && drag.kind !== 'colResize' && drag.kind !== 'rowResize') {
        autoScroll.current = requestAnimationFrame(stepAutoScroll);
      }
    };
    const up = () => endDrag();
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', up);
    return () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', up);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [drag?.kind]);

  // ─── указатель ─────────────────────────────────────────────────────────────
  const touchStart = useRef<{ x: number; y: number; zone: Zone } | null>(null);
  const pointInput = useRef<HTMLTextAreaElement | null>(null);
  /** Где и как нажали — чтобы отличить клик по ячейке артикула от протягивания, Shift/Ctrl и указания ссылки в формуле */
  const down = useRef<{ x: number; y: number; button: number; mods: boolean; pointMode: boolean } | null>(null);

  const onPointerDown = (e: React.PointerEvent) => {
    const target = e.target as HTMLElement;
    const ui = useUI.getState();
    down.current = { x: e.clientX, y: e.clientY, button: e.button, mods: e.shiftKey || e.ctrlKey || e.metaKey || e.altKey, pointMode: false };
    if (ui.menu) ui.set({ menu: null });

    // маркер скрытых строк/столбцов
    const unhide = target.closest<HTMLElement>('[data-unhide]');
    if (unhide) {
      e.preventDefault();
      unhideAt(unhide.dataset.unhide!, Number(unhide.dataset.at));
      return;
    }
    if (target.closest('.ch-filter') || target.closest('.g-add')) return;

    const resizeCol = target.closest<HTMLElement>('[data-resize-col]');
    if (resizeCol && e.button === 0) {
      e.preventDefault();
      const vc = Number(resizeCol.dataset.resizeCol);
      const w = geo.colX[vc + 1] - geo.colX[vc];
      setDrag({ kind: 'colResize', pointerId: e.pointerId, c: vc, startX: e.clientX, startW: w, w });
      return;
    }
    const resizeRow = target.closest<HTMLElement>('[data-resize-row]');
    if (resizeRow && e.button === 0) {
      e.preventDefault();
      const vr = Number(resizeRow.dataset.resizeRow);
      const h = geo.rowY[vr + 1] - geo.rowY[vr];
      setDrag({ kind: 'rowResize', pointerId: e.pointerId, vr, startY: e.clientY, startH: h, h });
      return;
    }
    if (target.closest('.fill-handle') && e.button === 0) {
      e.preventDefault();
      if (ui.edit && !commitEdit(null)) return;
      setDrag({ kind: 'fill', pointerId: e.pointerId, axis: null, to: 0 });
      return;
    }
    const openRow = target.closest<HTMLElement>('[data-open-row]');
    if (openRow) return;

    const zone = hit(e.clientX, e.clientY);
    if (e.pointerType === 'touch') {
      touchStart.current = { x: e.clientX, y: e.clientY, zone };
      return;
    }
    if (zone.z === 'none') return;
    const right = e.button === 2;

    // режим указания ссылки в формуле
    if (ui.edit && zone.z === 'cell' && !right) {
      const ae = document.activeElement as HTMLTextAreaElement | null;
      const ta = ae && ae.matches('.ed, .fbar-input') ? ae : null;
      pointInput.current = ta;
      const caret = ta?.selectionStart ?? ui.edit.text.length;
      if (isFormulaPointMode(ui.edit.text, caret) && !(zone.vr === ui.edit.r && zone.vc === ui.edit.c)) {
        e.preventDefault();
        down.current.pointMode = true;
        const d: Extract<Drag, { kind: 'point' }> = {
          kind: 'point',
          pointerId: e.pointerId,
          anchor: { vr: zone.vr, vc: zone.vc },
          start: caret,
          end: ta?.selectionEnd ?? caret,
        };
        setDrag(d);
        updatePointRef(d, zone.vr, zone.vc);
        return;
      }
    }
    if (ui.edit && !commitEdit(null)) return;
    if (!right) e.preventDefault();
    scroller.current?.focus({ preventScroll: true });
    applyZoneSelect(zone, e.shiftKey, right);
    if (!right) {
      if (zone.z === 'cell') setDrag({ kind: 'cells', pointerId: e.pointerId });
      else if (zone.z === 'col') setDrag({ kind: 'cols', pointerId: e.pointerId });
      else if (zone.z === 'row') setDrag({ kind: 'rows', pointerId: e.pointerId });
    }
  };

  const onPointerUp = (e: React.PointerEvent) => {
    if (e.pointerType !== 'touch' || !touchStart.current) return;
    const t = touchStart.current;
    touchStart.current = null;
    if (Math.hypot(e.clientX - t.x, e.clientY - t.y) > 8) return;
    const target = e.target as HTMLElement;
    if (target.closest('.c-link, [data-open-row]')) return;
    const ui = useUI.getState();
    if (ui.edit && !commitEdit(null)) return;
    applyZoneSelect(t.zone, false, false);
  };

  function applyZoneSelect(zone: Zone, shift: boolean, right: boolean) {
    const g = geoRef.current;
    const s = useUI.getState().sel;
    const { r1, r2, c1, c2 } = selRect(s);
    const lastR = g.phys.length - 1;
    const lastC = g.cols.length - 1;
    if (zone.z === 'cell') {
      if (right && zone.vr >= r1 && zone.vr <= r2 && zone.vc >= c1 && zone.vc <= c2) return;
      setSelection(shift ? { ...s, fr: zone.vr, fc: zone.vc } : { ar: zone.vr, ac: zone.vc, fr: zone.vr, fc: zone.vc });
    } else if (zone.z === 'col') {
      if (right && selectionIsWholeCols() && zone.vc >= c1 && zone.vc <= c2) return;
      if (shift) setSelection({ ar: 0, ac: s.ac, fr: lastR, fc: zone.vc }, false);
      else setSelection({ ar: 0, ac: zone.vc, fr: lastR, fc: zone.vc }, false);
    } else if (zone.z === 'row') {
      if (right && selectionIsWholeRows() && zone.vr >= r1 && zone.vr <= r2) return;
      if (shift) setSelection({ ar: s.ar, ac: 0, fr: zone.vr, fc: lastC }, false);
      else setSelection({ ar: zone.vr, ac: 0, fr: zone.vr, fc: lastC }, false);
    } else if (zone.z === 'corner') {
      setSelection({ ar: 0, ac: 0, fr: lastR, fc: lastC }, false);
    }
  }

  const onDoubleClick = (e: React.MouseEvent) => {
    const target = e.target as HTMLElement;
    // по ячейке артикула карточка уже открылась первым кликом
    if (target.closest('[data-open-card], .c-link, .ch-filter, [data-unhide], .g-add')) return;
    const rc = target.closest<HTMLElement>('[data-resize-col]');
    if (rc) {
      autofitColumn(Number(rc.dataset.resizeCol));
      return;
    }
    const rr = target.closest<HTMLElement>('[data-resize-row]');
    if (rr) {
      const rowId = geo.view.rows[Number(rr.dataset.resizeRow)];
      store.setRowsHeight(sheet, [rowId], undefined);
      return;
    }
    const zone = hit(e.clientX, e.clientY);
    if (zone.z === 'col') {
      useUI.getState().set({ dialog: { kind: 'rename-col', c: zone.vc } });
      return;
    }
    if (zone.z === 'row') {
      openCard(zone.vr);
      return;
    }
    if (zone.z !== 'cell') return;
    const x = cellAt(zone.vr, zone.vc);
    const d = x ? store.display(sheet, x.phys, x.c) : null;
    if (d?.img && !d.text) {
      useUI.getState().set({ lightbox: { imageId: d.img } });
      return;
    }
    startEdit('edit');
  };

  const onContextMenu = (e: React.MouseEvent) => {
    const target = e.target as HTMLElement;
    if (target.closest('.ed')) return;
    e.preventDefault();
    const zone = hit(e.clientX, e.clientY);
    if (zone.z === 'none' || zone.z === 'corner') return;
    if (e.nativeEvent instanceof PointerEvent && e.nativeEvent.pointerType === 'touch') applyZoneSelect(zone, false, true);
    const ui = useUI.getState();
    if (zone.z === 'col') ui.set({ menu: { kind: 'col', x: e.clientX, y: e.clientY, c: zone.vc } });
    else if (zone.z === 'row') ui.set({ menu: { kind: 'row', x: e.clientX, y: e.clientY, r: zone.vr } });
    else ui.set({ menu: { kind: 'cell', x: e.clientX, y: e.clientY } });
  };

  const onClick = (e: React.MouseEvent) => {
    const target = e.target as HTMLElement;
    const key = target.closest<HTMLElement>('[data-open-card]');
    if (key) {
      // ТЗ: карточка открывается по тексту артикула или по самой ячейке — но не при протягивании,
      // Shift/Ctrl-выделении и не когда мышью указывают ссылку в формуле
      const d = down.current;
      const moved = d ? Math.hypot(e.clientX - d.x, e.clientY - d.y) > 5 : false;
      if (d && (d.pointMode || d.mods || d.button !== 0 || moved)) return;
      if (useUI.getState().edit) return;
      openCard(Number(key.dataset.openCard));
      return;
    }
    const openRow = target.closest<HTMLElement>('[data-open-row]');
    if (openRow) openCard(Number(openRow.dataset.openRow));
  };

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (useUI.getState().edit) return;
    if (e.target !== scroller.current) return;
    handleGridKey(e);
  };

  // ─── перетаскивание файлов ─────────────────────────────────────────────────
  const onDragOver = (e: React.DragEvent) => {
    if (!Array.from(e.dataTransfer.types).includes('Files')) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    const z = hit(e.clientX, e.clientY, true);
    if (z.z === 'cell' && (dropTarget?.vr !== z.vr || dropTarget?.vc !== z.vc)) setDropTarget({ vr: z.vr, vc: z.vc });
  };
  const onDragLeave = (e: React.DragEvent) => {
    if (!scroller.current?.contains(e.relatedTarget as Node)) setDropTarget(null);
  };
  const onDrop = (e: React.DragEvent) => {
    const files = Array.from(e.dataTransfer.files);
    setDropTarget(null);
    if (!files.length) return;
    e.preventDefault();
    const sheetFile = files.find((f) => /\.(xlsx|csv)$/i.test(f.name));
    if (sheetFile) {
      window.dispatchEvent(new CustomEvent('reestr:import-file', { detail: sheetFile }));
      return;
    }
    const z = hit(e.clientX, e.clientY, true);
    if (z.z !== 'cell') return;
    const images = files.filter(isImageFile);
    if (!images.length) {
      useUI.getState().toast({ text: 'Перетащите фото (JPG, PNG, WebP) или файл .xlsx', tone: 'error' });
      return;
    }
    setSelection({ ar: z.vr, ac: z.vc, fr: z.vr, fc: z.vc }, false);
    void insertImages(images, z.vr, z.vc);
  };

  // ─── видимые диапазоны ─────────────────────────────────────────────────────
  const n = geo.phys.length;
  const bodyVisH = Math.max(0, vp.h - HD_H);
  const r0 = n ? Math.max(0, indexAt(geo.rowY, vp.top) - OVERSCAN_ROWS) : 0;
  const r1 = n ? Math.min(n - 1, indexAt(geo.rowY, vp.top + bodyVisH) + OVERSCAN_ROWS) : -1;
  const mainVisW = Math.max(0, vp.w - leftW);
  const nc = geo.cols.length;
  const c0 = nc > geo.frozen ? Math.max(geo.frozen, indexAt(geo.colX, vp.left + geo.frozenW) - 1) : geo.frozen;
  const c1 = nc > geo.frozen ? Math.min(nc - 1, indexAt(geo.colX, vp.left + geo.frozenW + mainVisW) + 1) : geo.frozen - 1;

  const keyCol = sheet.keyColId ? store.colIndexOf(sheet, sheet.keyColId) : -1;
  const sheetHasColFormula = sheet.columns.some((c) => c.formula);

  const rows = [];
  const leftRows = [];
  const rowHeads = [];
  const { r1: sr1, r2: sr2, c1: sc1, c2: sc2 } = selRect(sel);
  const wholeRows = sc1 === 0 && sc2 === nc - 1;
  const wholeCols = sr1 === 0 && sr2 === n - 1 && n > 0;
  for (let vr = r0; vr <= r1; vr++) {
    const rowId = geo.view.rows[vr];
    const row = sheet.rows.get(rowId);
    const y = geo.rowY[vr];
    const h = geo.rowY[vr + 1] - y;
    const phys = geo.phys[vr];
    let hasF = sheetHasColFormula;
    if (!hasF && row) for (const k in row.cells) if (row.cells[k].f) {
      hasF = true;
      break;
    }
    const calcKey = hasF ? store.calcEpoch : 0;
    if (c1 >= c0)
      rows.push(
        <GridRow
          key={rowId}
          geo={geo}
          sheet={sheet}
          row={row}
          vr={vr}
          phys={phys}
          y={y}
          h={h}
          c0={c0}
          c1={c1}
          xShift={geo.frozenW}
          keyCol={keyCol}
          calcKey={calcKey}
          layoutKey={store.layoutVersion}
        />,
      );
    if (geo.frozen > 0)
      leftRows.push(
        <GridRow
          key={rowId}
          geo={geo}
          sheet={sheet}
          row={row}
          vr={vr}
          phys={phys}
          y={y}
          h={h}
          c0={0}
          c1={geo.frozen - 1}
          xShift={0}
          keyCol={keyCol}
          calcKey={calcKey}
          layoutKey={store.layoutVersion}
        />,
      );
    const selected = vr >= sr1 && vr <= sr2;
    rowHeads.push(
      <div
        key={rowId}
        className={'rh' + (selected ? (wholeRows ? ' is-full' : ' is-sel') : '')}
        style={{ transform: `translateY(${y}px)`, height: h }}
      >
        {geo.hiddenBefore.has(vr) && (
          <button type="button" className="hid-mark hid-mark--row" data-unhide="row" data-at={vr} aria-label="Показать скрытые строки" />
        )}
        <span className="rh-n">{phys + 1}</span>
        <button type="button" className="rh-open" data-open-row={vr} tabIndex={-1} aria-label={`Открыть карточку строки ${phys + 1}`}>
          <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden>
            <path d="M4.5 2.5h5v5M9.5 2.5 3 9" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
        <div className="rh-resize" data-resize-row={vr} />
      </div>,
    );
  }

  // ─── заголовки столбцов ────────────────────────────────────────────────────
  const colHead = (vc: number, xShift: number) => {
    const c = geo.cols[vc];
    const col = sheet.columns[c];
    const x = geo.colX[vc] - xShift;
    const w = geo.colX[vc + 1] - geo.colX[vc];
    const filtered = !!sheet.filters[col.id];
    const selected = vc >= sc1 && vc <= sc2;
    const isKey = c === keyCol;
    const open = filterMenu?.c === vc;
    return (
      <div
        key={col.id}
        className={
          'ch' +
          (selected ? (wholeCols ? ' is-full' : ' is-sel') : '') +
          (filtered ? ' is-filtered' : '') +
          (isKey ? ' is-key' : '') +
          (col.formula ? ' is-formula' : '')
        }
        style={{ left: x, width: w }}
        title={
          isKey
            ? 'Клик по ячейке открывает карточку товара. Изменить значение — выделите ячейку стрелками и нажмите F2 или правьте в строке формул'
            : col.formula
              ? `Формула столбца: ${col.formula}`
              : undefined
        }
      >
        {geo.hiddenColBefore.has(vc) && (
          <button type="button" className="hid-mark hid-mark--col" data-unhide="col" data-at={vc} aria-label="Показать скрытые столбцы" />
        )}
        <span className="ch-l">{colToLetters(c)}</span>
        {isKey && <KeyGlyph />}
        <span className="ch-n">{col.name}</span>
        <button
          type="button"
          className={'ch-filter' + (open ? ' is-open' : '')}
          aria-label={`Фильтр и сортировка: ${col.name}`}
          aria-haspopup="dialog"
          aria-expanded={open}
          tabIndex={-1}
          onClick={(e) => {
            const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
            const ui = useUI.getState();
            if (ui.edit && !commitEdit(null)) return;
            ui.set({ filterMenu: open ? null : { c: vc, rect: { left: r.left, top: r.top, bottom: r.bottom, right: r.right } }, menu: null });
          }}
        >
          <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden>
            {filtered ? (
              <path d="M1 1.5h8L6 5.2V9L4 8V5.2Z" fill="currentColor" />
            ) : (
              <path d="M2 3.5 5 6.5 8 3.5" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
            )}
          </svg>
        </button>
        <div className="ch-resize" data-resize-col={vc} />
      </div>
    );
  };

  const headsMain = [];
  for (let vc = c0; vc <= c1; vc++) headsMain.push(colHead(vc, geo.frozenW));
  const headsLeft = [];
  for (let vc = 0; vc < geo.frozen; vc++) headsLeft.push(colHead(vc, -RH_W));

  // ─── слои выделения ────────────────────────────────────────────────────────
  const paneLayers = (pane: 'left' | 'main') => {
    const pStart = pane === 'left' ? 0 : geo.frozen;
    const pEnd = pane === 'left' ? geo.frozen - 1 : nc - 1;
    const xShift = pane === 'left' ? -RH_W : geo.frozenW;
    const out: React.ReactNode[] = [];
    if (!n || pEnd < pStart) return out;
    const rect = (rr1: number, rr2: number, cc1: number, cc2: number) => {
      const a = Math.max(cc1, pStart);
      const b = Math.min(cc2, pEnd);
      if (a > b) return null;
      return {
        left: geo.colX[a] - xShift,
        width: geo.colX[b + 1] - geo.colX[a],
        top: geo.rowY[rr1],
        height: geo.rowY[rr2 + 1] - geo.rowY[rr1],
        openL: cc1 < pStart,
        openR: cc2 > pEnd,
      };
    };
    const multi = sr1 !== sr2 || sc1 !== sc2;
    const rg = rect(sr1, Math.min(sr2, n - 1), sc1, sc2);
    if (rg && multi)
      out.push(
        <div
          key="range"
          className={'sel-range' + (rg.openL ? ' open-l' : '') + (rg.openR ? ' open-r' : '')}
          style={{ left: rg.left - 1, top: rg.top - 1, width: rg.width + 1, height: rg.height + 1 }}
        />,
      );
    if (sel.ac >= pStart && sel.ac <= pEnd && sel.ar < n && !edit) {
      const a = rect(sel.ar, sel.ar, sel.ac, sel.ac)!;
      out.push(<div key="active" className="sel-active" style={{ left: a.left - 1, top: a.top - 1, width: a.width + 1, height: a.height + 1 }} />);
    }
    if (rg && sc2 >= pStart && sc2 <= pEnd && !edit) {
      out.push(
        <div
          key="fill"
          className="fill-handle"
          style={{ left: rg.left + rg.width - 5, top: rg.top + rg.height - 5 }}
          aria-hidden
        />,
      );
    }
    if (drag?.kind === 'fill' && drag.axis) {
      const f =
        drag.axis === 'row'
          ? rect(Math.min(sr1, drag.to), Math.max(sr2, drag.to), sc1, sc2)
          : rect(sr1, sr2, Math.min(sc1, drag.to), Math.max(sc2, drag.to));
      if (f) out.push(<div key="fillp" className="fill-preview" style={{ left: f.left, top: f.top, width: f.width, height: f.height }} />);
    }
    if (copyRange) {
      const cr = rect(Math.min(copyRange.r1, n - 1), Math.min(copyRange.r2, n - 1), copyRange.c1, copyRange.c2);
      if (cr) out.push(<div key="copy" className="copy-range" style={{ left: cr.left, top: cr.top, width: cr.width, height: cr.height }} />);
    }
    if (dropTarget) {
      const dt = rect(dropTarget.vr, dropTarget.vr, dropTarget.vc, dropTarget.vc);
      if (dt) out.push(<div key="drop" className="drop-target" style={{ left: dt.left, top: dt.top, width: dt.width, height: dt.height }} />);
    }
    // ссылки редактируемой формулы
    if (edit && edit.text.startsWith('=')) {
      listRefs(edit.text).forEach((ref, i) => {
        if (ref.sheet && ref.sheet.toLocaleLowerCase('ru') !== sheet.name.toLocaleLowerCase('ru')) return;
        const pr1 = ref.a.r ?? 0;
        const pr2 = ref.b ? (ref.b.r ?? sheet.rowOrder.length - 1) : pr1;
        const pc1 = ref.a.c ?? 0;
        const pc2 = ref.b ? (ref.b.c ?? sheet.columns.length - 1) : pc1;
        const vr1 = geo.vrOf[Math.min(pr1, pr2)] ?? -1;
        const vr2 = geo.vrOf[Math.min(Math.max(pr1, pr2), geo.vrOf.length - 1)] ?? -1;
        const vc1 = geo.vcOf[Math.min(pc1, pc2)] ?? -1;
        const vc2 = geo.vcOf[Math.min(Math.max(pc1, pc2), geo.vcOf.length - 1)] ?? -1;
        if (vr1 < 0 || vr2 < 0 || vc1 < 0 || vc2 < 0) return;
        const rr = rect(vr1, vr2, vc1, vc2);
        if (!rr) return;
        const color = REF_COLORS[i % REF_COLORS.length];
        out.push(
          <div
            key={'ref' + i}
            className="ref-box"
            style={{ left: rr.left, top: rr.top, width: rr.width, height: rr.height, borderColor: color, backgroundColor: color + '14' }}
          />,
        );
      });
    }
    if (drag?.kind === 'colResize' && drag.c >= pStart && drag.c <= pEnd) {
      out.push(<div key="guide" className="resize-guide resize-guide--v" style={{ left: geo.colX[drag.c] - xShift + drag.w - 1 }} />);
    }
    if (drag?.kind === 'rowResize') {
      out.push(<div key="guide" className="resize-guide resize-guide--h" style={{ top: geo.rowY[drag.vr] + drag.h - 1 }} />);
    }
    if (edit && edit.c >= pStart && edit.c <= pEnd) {
      out.push(<CellEditor key="editor" geo={geo} xShift={xShift} pane={pane} />);
    }
    return out;
  };

  // ─── пусто ─────────────────────────────────────────────────────────────────
  const emptyState =
    n === 0 ? (
      <div className="g-empty" style={{ left: 0, top: HD_H + 24 }}>
        {geo.view.filtered ? (
          <>
            <p>Под условия фильтра не подходит ни одна строка.</p>
            <button
              type="button"
              className="btn"
              onClick={() => {
                store.clearFilters(sheet);
                store.setSearch('');
              }}
            >
              Сбросить фильтры и поиск
            </button>
          </>
        ) : (
          <>
            <p>Все строки скрыты.</p>
            <button type="button" className="btn" onClick={() => store.showAllRows(sheet)}>
              Показать все строки
            </button>
          </>
        )}
      </div>
    ) : null;

  const addRowBtn = (
    <button
      type="button"
      className="g-add g-add--row"
      style={{ top: geo.totalH, width: RH_W + Math.min(geo.totalW, 260) }}
      onClick={() => {
        const vr = appendRow();
        if (vr >= 0) setSelection({ ar: vr, ac: 0, fr: vr, fc: 0 });
        scroller.current?.focus({ preventScroll: true });
      }}
    >
      <span aria-hidden>+</span> Строка
    </button>
  );

  const unhideAt = (kind: string, at: number) => {
    const g = geoRef.current;
    if (kind === 'row') {
      const to = at < g.phys.length ? g.phys[at] : sheet.rowOrder.length;
      const from = at > 0 ? g.phys[at - 1] + 1 : 0;
      const ids = sheet.rowOrder.slice(from, to).filter((id) => sheet.rows.get(id)?.hidden);
      store.setRowsHidden(sheet, ids, false);
    } else {
      const to = at < g.cols.length ? g.cols[at] : sheet.columns.length;
      const from = at > 0 ? g.cols[at - 1] + 1 : 0;
      store.setColumnsHidden(
        sheet,
        sheet.columns.slice(from, to).filter((c) => c.hidden).map((c) => c.id),
        false,
      );
    }
  };

  function autofitColumn(vc: number) {
    const c = geo.cols[vc];
    const col = sheet.columns[c];
    const canvas = document.createElement('canvas');
    const cx = canvas.getContext('2d')!;
    cx.font = '13px "Golos Text Variable", system-ui, sans-serif';
    let w = cx.measureText(col.name).width + 64;
    const limit = Math.min(sheet.rowOrder.length, 3000);
    for (let i = 0; i < limit; i++) {
      const d = store.display(sheet, i, c);
      if (d.text) w = Math.max(w, cx.measureText(d.text).width * (d.style.b ? 1.06 : 1) + 22);
    }
    store.updateColumn(sheet, col.id, { w: Math.round(Math.max(48, Math.min(560, w))) }, 'Ширина по содержимому');
  }

  return (
    <div
      ref={scroller}
      className={'g' + (drag?.kind === 'colResize' ? ' is-col-resizing' : '') + (drag?.kind === 'rowResize' ? ' is-row-resizing' : '')}
      tabIndex={0}
      role="grid"
      aria-label={`Лист «${sheet.name}»`}
      aria-rowcount={n}
      aria-colcount={nc}
      aria-activedescendant={undefined}
      onScroll={onScroll}
      onPointerDown={onPointerDown}
      onPointerUp={onPointerUp}
      onDoubleClick={onDoubleClick}
      onContextMenu={onContextMenu}
      onClick={onClick}
      onKeyDown={onKeyDown}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <div className="g-inner" style={{ width: leftW + mainW, height: HD_H + geo.totalH + ADD_ROW_H }}>
        <div className="g-head" style={{ width: leftW + mainW }}>
          <div className="g-corner" style={{ width: leftW }}>
            <button
              type="button"
              className="g-corner-box"
              tabIndex={-1}
              aria-label="Выделить всё"
              onClick={() => setSelection({ ar: 0, ac: 0, fr: n - 1, fc: nc - 1 }, false)}
            />
            {headsLeft}
          </div>
          <div className="g-heads" style={{ width: mainW }}>
            {headsMain}
            {geo.hiddenAfterCols && (
              <button
                type="button"
                className="hid-mark hid-mark--col"
                data-unhide="col"
                data-at={nc}
                style={{ left: geo.totalW - geo.frozenW }}
                aria-label="Показать скрытые столбцы"
              />
            )}
            <button
              type="button"
              className="g-add g-add--col"
              style={{ left: geo.totalW - geo.frozenW }}
              aria-label="Добавить столбец"
              title="Добавить столбец"
              onClick={() => {
                store.insertColumns(sheet, sheet.columns.length, 1);
                const v = store.view(sheet);
                setSelection({ ar: 0, ac: v.cols.length - 1, fr: v.rows.length - 1, fc: v.cols.length - 1 });
              }}
            >
              +
            </button>
          </div>
        </div>
        <div className="g-body" style={{ height: geo.totalH + ADD_ROW_H }}>
          <div className="g-left" style={{ width: leftW }}>
            <div className="g-rh">{rowHeads}</div>
            {leftRows}
            {paneLayers('left')}
            {geo.hiddenAfterRows && (
              <button
                type="button"
                className="hid-mark hid-mark--row"
                data-unhide="row"
                data-at={n}
                style={{ top: geo.totalH }}
                aria-label="Показать скрытые строки"
              />
            )}
            {!geo.view.filtered && addRowBtn}
          </div>
          <div className="g-main" style={{ width: mainW }}>
            {rows}
            {paneLayers('main')}
          </div>
        </div>
        {emptyState}
      </div>
    </div>
  );
}

/** Знак ключевого столбца: крошечный штрихкод — клик по значению открывает карточку. */
function KeyGlyph() {
  return (
    <svg className="ch-key" width="11" height="10" viewBox="0 0 11 10" aria-label="Ключевой столбец карточки" role="img">
      <path d="M0 0h1.5v10H0zM3 0h1v10H3zM5.5 0H8v10H5.5zM9.3 0h.7v10h-.7z" fill="currentColor" />
    </svg>
  );
}

function selectionIsWholeCols() {
  const { view } = ctx();
  const { r1, r2 } = selRect(useUI.getState().sel);
  return r1 === 0 && r2 === view.rows.length - 1;
}

function selectionIsWholeRows() {
  const { view } = ctx();
  const { c1, c2 } = selRect(useUI.getState().sel);
  return c1 === 0 && c2 === view.cols.length - 1;
}

export type { Geometry };
