import type { ComplicationArtId } from '../content/slides';

const WALL = '#c47e92', LUMEN = '#f6d4db', BURG = '#7a1f3d', RED = '#a3203f', INK = '#8d6a77';

export default function ComplicationArt({ id }: { id: ComplicationArtId }) {
  const wall = { stroke: WALL, strokeWidth: 5, fill: 'none', strokeLinecap: 'round' as const };
  return (
    <svg viewBox="0 0 160 110" className="comp-svg" role="img" aria-label={id}>
      {id === 'obstruction' && (
        <g>
          <path d="M10 28 H85 Q100 28 108 40 H150 V70 H108 Q100 82 85 82 H10 Z" fill={LUMEN} />
          <path d="M10 28 H85 Q100 28 108 40 H150" {...wall} /><path d="M10 82 H85 Q100 82 108 70 H150" {...wall} />
          <ellipse cx="113" cy="55" rx="13" ry="15" fill={BURG} />
          {[30, 52, 70].map((x, i) => <circle key={x} cx={x} cy={44 + i * 10} r="6" fill="#e8b2bf" />)}
          <path d="M20 100 H80" stroke={INK} strokeWidth="2" strokeDasharray="4 4" /><path d="M80 100 l-6 -4 v8z" fill={INK} />
        </g>
      )}
      {id === 'bleeding' && (
        <g>
          <rect x="10" y="32" width="140" height="46" fill={LUMEN} />
          <path d="M10 32 H150" {...wall} /><path d="M10 78 H150" {...wall} />
          <ellipse cx="80" cy="34" rx="20" ry="12" fill={BURG} />
          {[[80, 56], [90, 68], [72, 70]].map(([x, y]) => <path key={x} d={`M${x} ${y - 7} q5 7 0 11 q-5 -4 0 -11z`} fill={RED} />)}
        </g>
      )}
      {id === 'perforation' && (
        <g>
          <rect x="10" y="32" width="140" height="46" fill={LUMEN} />
          <path d="M10 32 H150" {...wall} /><path d="M10 78 H68" {...wall} /><path d="M92 78 H150" {...wall} />
          {[[80, 90], [74, 98], [88, 100], [82, 106]].map(([x, y]) => <circle key={`${x}${y}`} cx={x} cy={y} r="3.5" fill={RED} opacity="0.8" />)}
          <path d="M80 60 V84" stroke={RED} strokeWidth="2" strokeDasharray="3 3" />
        </g>
      )}
      {id === 'metastasis' && (
        <g>
          <defs><marker id="ah-met" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto"><path d="M0 0 L10 5 L0 10z" fill={BURG} /></marker></defs>
          <path d="M16 90 V28 H52 V90" stroke="#e29bad" strokeWidth="9" fill="none" strokeLinecap="round" strokeLinejoin="round" />
          <circle cx="52" cy="60" r="6" fill={BURG} />
          <path d="M96 20 C120 8 150 16 152 32 C150 48 120 44 104 40 C94 36 90 28 96 20Z" fill="#c9605f" />
          <ellipse cx="108" cy="84" rx="11" ry="17" fill="#e9a8b8" /><ellipse cx="134" cy="84" rx="11" ry="17" fill="#e9a8b8" />
          <path d="M60 56 L92 36" stroke={BURG} strokeWidth="2" strokeDasharray="4 3" markerEnd="url(#ah-met)" />
          <path d="M60 64 L96 82" stroke={BURG} strokeWidth="2" strokeDasharray="4 3" markerEnd="url(#ah-met)" />
        </g>
      )}
      {id === 'malnutrition' && (
        <g fill="none">
          <circle cx="46" cy="58" r="32" fill="#fff" stroke={WALL} strokeWidth="4" />
          <circle cx="46" cy="58" r="20" stroke="#e8b2bf" strokeWidth="2" />
          <path d="M96 30 V86 H152" stroke={INK} strokeWidth="2" />
          <path d="M102 36 L128 52 L146 80" stroke={BURG} strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round" />
          <path d="M146 80 l-9 -2 l5 -8z" fill={BURG} stroke="none" />
        </g>
      )}
      {id === 'emotional' && (
        <g fill="none">
          <circle cx="58" cy="42" r="22" fill="#f6d4db" stroke={WALL} strokeWidth="4" />
          <path d="M14 104 Q16 72 58 70 Q100 72 102 104" stroke={WALL} strokeWidth="4" strokeLinecap="round" />
          <path d="M112 30 q10 -14 20 0 t20 0" stroke={BURG} strokeWidth="3" strokeLinecap="round" />
          <path d="M112 48 q10 -14 20 0 t20 0" stroke="#c9567b" strokeWidth="3" strokeLinecap="round" />
          <path d="M112 66 q10 -14 20 0 t20 0" stroke="#e29bad" strokeWidth="3" strokeLinecap="round" />
        </g>
      )}
    </svg>
  );
}
