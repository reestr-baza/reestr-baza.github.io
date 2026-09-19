import { useState } from 'react';

const NEUTRALS = ['#ffffff', '#f2f2f2', '#d9d9d9', '#a6a6a6', '#595959', '#000000'];
const TINTS = ['#fde2e1', '#feebd6', '#fff6c7', '#dff3e4', '#d8f1f0', '#dce9fb', '#ece3fa', '#fbe3f0'];
const MIDS = ['#f6a5a0', '#fbc48f', '#ffe380', '#a6ddb4', '#93d6d2', '#a9c8f5', '#c9b3f0', '#f2a9ce'];
const STRONG = ['#d93025', '#e8710a', '#c99a00', '#1e8e3e', '#12877e', '#1a5fd1', '#7b3fd1', '#c2185b'];
/** Стандартные цвета Excel — чтобы привычные заливки находились там же */
const EXCEL = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];

const NAMES: Record<string, string> = {
  '#ffffff': 'белый',
  '#000000': 'чёрный',
  '#ffff00': 'жёлтый (Excel)',
  '#ff0000': 'красный (Excel)',
  '#00b050': 'зелёный (Excel)',
  '#0070c0': 'синий (Excel)',
};

let recent: string[] = [];

export function ColorPicker({
  value,
  onPick,
  noneLabel,
}: {
  value: string | undefined;
  onPick: (c: string | undefined) => void;
  noneLabel: string;
}) {
  const [custom, setCustom] = useState(value ?? '#1a5fd1');
  const swatch = (c: string) => (
    <button
      key={c}
      type="button"
      className={'sw' + (value?.toLowerCase() === c ? ' is-on' : '')}
      style={{ backgroundColor: c }}
      aria-label={NAMES[c] ?? c}
      title={NAMES[c] ?? c}
      onClick={() => onPick(c)}
    />
  );
  return (
    <div className="cp">
      <button type="button" className="cp-none" onClick={() => onPick(undefined)}>
        <span className="cp-none-sw" aria-hidden />
        {noneLabel}
      </button>
      <div className="cp-grid cp-grid--8">
        {NEUTRALS.map(swatch)}
        <span />
        <span />
        {TINTS.map(swatch)}
        {MIDS.map(swatch)}
        {STRONG.map(swatch)}
      </div>
      <div className="cp-label">Стандартные Excel</div>
      <div className="cp-grid cp-grid--10">{EXCEL.map(swatch)}</div>
      {recent.length > 0 && (
        <>
          <div className="cp-label">Недавние</div>
          <div className="cp-grid cp-grid--10">{recent.map(swatch)}</div>
        </>
      )}
      <label className="cp-custom">
        <input
          type="color"
          value={custom}
          onChange={(e) => setCustom(e.target.value)}
          aria-label="Свой цвет"
        />
        <span>Свой цвет</span>
        <button
          type="button"
          className="btn btn--sm"
          onClick={() => {
            recent = [custom, ...recent.filter((c) => c !== custom)].slice(0, 10);
            onPick(custom);
          }}
        >
          Применить
        </button>
      </label>
    </div>
  );
}
