/** Иконки, которых нет в lucide: вертикальное выравнивание и знак строки-этикетки. Тот же штрих, что у lucide. */
const base = {
  width: 16,
  height: 16,
  viewBox: '0 0 24 24',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 1.75,
  strokeLinecap: 'round' as const,
  strokeLinejoin: 'round' as const,
  'aria-hidden': true,
};

export const VAlignTop = () => (
  <svg {...base}>
    <path d="M4 4h16" />
    <path d="M9 9h6M9 13h6" />
    <path d="M12 17v3" opacity=".35" />
  </svg>
);

export const VAlignMiddle = () => (
  <svg {...base}>
    <path d="M4 4h16M4 20h16" opacity=".35" />
    <path d="M9 10h6M9 14h6" />
  </svg>
);

export const VAlignBottom = () => (
  <svg {...base}>
    <path d="M4 20h16" />
    <path d="M9 11h6M9 15h6" />
    <path d="M12 4v3" opacity=".35" />
  </svg>
);

/** Высота строк: три полосы разной высоты */
export const RowsDensity = ({ level }: { level: 'S' | 'M' | 'L' }) => (
  <svg {...base}>
    {level === 'S' && <path d="M4 6h16M4 10h16M4 14h16M4 18h16" />}
    {level === 'M' && <path d="M4 5h16M4 12h16M4 19h16" />}
    {level === 'L' && <path d="M4 4h16M4 20h16" />}
  </svg>
);

/** Марка «Реестра»: штрихкод из пяти полос — то, чем помечен каждый товар на складе. */
export const BrandMark = ({ size = 22 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 22 22" aria-hidden>
    <rect x="0" y="0" width="22" height="22" rx="3" fill="var(--tape)" />
    <g fill="var(--ink)">
      <rect x="4" y="5" width="2" height="12" />
      <rect x="7.5" y="5" width="1" height="12" />
      <rect x="10" y="5" width="3" height="12" />
      <rect x="14.5" y="5" width="1" height="12" />
      <rect x="16.5" y="5" width="1.5" height="12" />
    </g>
  </svg>
);
