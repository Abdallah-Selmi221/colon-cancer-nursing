import type { SurgeryArtId } from '../content/slides';

const BURG = '#7a1f3d';
const COLON = 'M30 112 V34 H150 V112 M150 112 C150 130 120 118 110 132 S96 140 94 150';

export function SurgeryArt({ id }: { id: SurgeryArtId }) {
  return (
    <svg viewBox="0 0 180 170" className="surg-svg" role="img" aria-label={id}>
      {id === 'polypectomy' && (
        <g>
          <ellipse cx="90" cy="86" rx="76" ry="62" fill="#f6d4db" stroke="#c47e92" strokeWidth="5" />
          <path d="M90 148 V118" stroke="#d77a92" strokeWidth="9" strokeLinecap="round" />
          <circle cx="90" cy="108" r="19" fill="#d4627a" />
          <ellipse cx="90" cy="102" rx="25" ry="19" fill="none" stroke="#5c4a52" strokeWidth="2.5" />
          <path d="M115 98 L168 44" stroke="#5c4a52" strokeWidth="5" strokeLinecap="round" />
        </g>
      )}
      {id === 'colectomy' && (
        <g fill="none" strokeLinecap="round" strokeLinejoin="round">
          <path d={COLON} stroke="#e29bad" strokeWidth="16" />
          <path d="M150 52 V96" stroke={BURG} strokeWidth="16" />
          <path d="M136 46 H164 M136 102 H164" stroke="#5c4a52" strokeWidth="2.5" strokeDasharray="4 3" />
        </g>
      )}
      {id === 'multivisceral' && (
        <g fill="none" strokeLinecap="round" strokeLinejoin="round">
          <path d={COLON} stroke="#e29bad" strokeWidth="14" />
          <path d="M150 54 V94" stroke={BURG} strokeWidth="14" />
          <path d="M118 22 C138 14 170 20 172 36 C170 52 140 46 124 42 C114 38 112 28 118 22Z" fill="#c9605f" transform="translate(-18 6)" />
          <ellipse cx="136" cy="62" rx="40" ry="52" stroke="#5c4a52" strokeWidth="2.5" strokeDasharray="5 4" />
        </g>
      )}
    </svg>
  );
}

export function TherapyArt({ id }: { id: 'chemo' | 'radiation' | 'targeted' }) {
  return (
    <svg viewBox="0 0 120 120" className="therapy-svg" role="img" aria-label={id}>
      {id === 'chemo' && (
        <g>
          <rect x="38" y="8" width="44" height="64" rx="12" fill="#f6d4db" stroke="#c47e92" strokeWidth="3" />
          <rect x="43" y="34" width="34" height="33" rx="8" fill="#a8c5e0" />
          <path d="M60 72 V96 M60 96 q0 14 -12 18" stroke="#8d6a77" strokeWidth="3" fill="none" strokeLinecap="round" />
          <path d="M60 76 q-4 6 0 9 q4 -3 0 -9z" fill="#7aa5cf" />
        </g>
      )}
      {id === 'radiation' && (
        <g fill="none" strokeLinecap="round">
          <circle cx="60" cy="86" r="14" fill="#c9567b" stroke="none" />
          <path d="M20 12 L52 74 M60 8 V70 M100 12 L68 74" stroke="#a56f80" strokeWidth="4" strokeDasharray="1 9" />
          <rect x="34" y="4" width="52" height="12" rx="6" fill="#d9b8c4" stroke="none" />
        </g>
      )}
      {id === 'targeted' && (
        <g fill="none">
          <circle cx="56" cy="64" r="44" fill="#fff" stroke="#c47e92" strokeWidth="4" />
          <circle cx="56" cy="64" r="28" stroke="#c9567b" strokeWidth="4" />
          <circle cx="56" cy="64" r="11" fill="#7a1f3d" stroke="none" />
          <path d="M112 8 L60 60" stroke="#5c4a52" strokeWidth="4" strokeLinecap="round" />
          <path d="M112 8 l-18 2 M112 8 l-2 18" stroke="#5c4a52" strokeWidth="4" strokeLinecap="round" />
        </g>
      )}
    </svg>
  );
}

export function StomaArt({ mode }: { mode: 'site' | 'stoma' }) {
  return mode === 'site' ? (
    <g>
      <circle cx="96" cy="150" r="20" fill="none" stroke="#7a1f3d" strokeWidth="3" strokeDasharray="6 5" />
      <path d="M86 150 H106 M96 140 V160" stroke="#7a1f3d" strokeWidth="3" strokeLinecap="round" />
    </g>
  ) : (
    <g>
      <circle cx="96" cy="150" r="17" fill="#f3a3b2" />
      <circle cx="96" cy="150" r="10" fill="#d9566f" />
      <path d="M72 150 H120 V206 Q96 216 72 206 Z" fill="#fff" fillOpacity="0.75" stroke="#a56f80" strokeWidth="2.5" />
    </g>
  );
}
