const TUMOR = '#8a2145', POLYP = '#d77a92';

export function MetastasisArt() {
  return (
    <>
      <path d="M150 20 C176 8 206 16 208 34 C206 50 176 46 160 42 C150 38 146 28 150 20Z" fill="#c9605f" />
      <ellipse cx="170" cy="96" rx="12" ry="19" fill="#e9a8b8" /><ellipse cx="198" cy="96" rx="12" ry="19" fill="#e9a8b8" />
      <circle cx="40" cy="68" r="9" fill={TUMOR} />
      <path d="M52 62 L142 36 M52 74 L156 92" stroke={TUMOR} strokeWidth="2" strokeDasharray="5 4" fill="none" />
    </>
  );
}

/** Cross-section of the colon wall with a tumor that deepens from stage 0 to IV. */
export default function StageArt({ stage }: { stage: string }) {
  return (
    <svg viewBox="0 0 240 140" className="stage-svg" role="img" aria-label={`Stage ${stage} diagram`}>
      {stage !== 'IV' ? (
        <g>
          <rect x="0" y="0" width="240" height="40" fill="#fbe9ed" />
          <rect x="0" y="40" width="240" height="20" fill="#e8a3b3" />
          <rect x="0" y="60" width="240" height="28" fill="#d9879b" />
          <rect x="0" y="88" width="240" height="18" fill="#c9b0b8" />
          {stage === '0' && <><ellipse cx="120" cy="40" rx="15" ry="13" fill={POLYP} />{[[114, 38], [122, 36], [118, 44]].map(([x, y]) => <circle key={x} cx={x} cy={y} r="2.2" fill="#f6c3cd" />)}</>}
          {stage === 'I' && <ellipse cx="120" cy="50" rx="24" ry="20" fill={TUMOR} />}
          {stage === 'II' && <ellipse cx="120" cy="62" rx="34" ry="32" fill={TUMOR} />}
          {stage === 'III' && (
            <>
              <ellipse cx="120" cy="66" rx="34" ry="38" fill={TUMOR} />
              {[[58, 122], [122, 126], [188, 120]].map(([x, y]) => <ellipse key={x} cx={x} cy={y} rx="11" ry="8" fill="#c9567b" />)}
              <path d="M120 100 L122 118 M100 100 L66 116 M140 100 L180 114" stroke="#c9567b" strokeWidth="2" strokeDasharray="3 3" />
            </>
          )}
        </g>
      ) : (
        <g><MetastasisArt /></g>
      )}
    </svg>
  );
}
