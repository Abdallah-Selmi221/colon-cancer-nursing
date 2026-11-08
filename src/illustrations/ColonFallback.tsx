import type { SectionKey } from '../content/slides';

const PATHS: Record<SectionKey, string> = {
  ascending: 'M-1.6 -1.2 L-1.62 1',
  transverse: 'M-1.6 1 C-1.1 1.1 -0.6 0.5 0 0.5 S1.1 1.1 1.6 1.1',
  descending: 'M1.6 1.1 L1.6 -0.95',
  sigmoid: 'M1.6 -0.95 C1.3 -1.4 0.9 -0.9 0.5 -1.4 S0.1 -1.6 0.05 -1.95',
};

/** Flat colon drawing used when WebGL is unavailable. Same section selection logic as the 3D model. */
export default function ColonFallback({ active = null, tumor = false }: { active?: SectionKey | null; tumor?: boolean }) {
  return (
    <svg viewBox="-2.6 -2.9 5.2 5.2" className="fallback-svg" role="img" aria-label="Colon diagram">
      <g transform="scale(1,-1)" strokeLinecap="round" fill="none" strokeWidth={0.5}>
        {(Object.keys(PATHS) as SectionKey[]).map((k) => (
          <path key={k} d={PATHS[k]} stroke={active === null ? '#d98ba0' : active === k ? '#8a2145' : '#ecc7d0'} />
        ))}
        <path d="M0.05 -1.95 L0.02 -2.4" stroke="#e5b3bf" strokeWidth={0.4} />
        {tumor && <circle cx={1.85} cy={-0.3} r={0.22} fill="#7a1f3d" stroke="none" />}
      </g>
    </svg>
  );
}
