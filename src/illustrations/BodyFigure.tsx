import type { SpotId } from '../content/slides';

const BODY =
  'M78 80 Q110 92 142 80 L170 96 Q188 106 190 130 L196 232 Q197 244 188 244 Q180 244 179 232 L172 150 L168 258 L150 262 L140 420 Q139 430 128 430 L112 430 L111 300 L109 300 L108 430 L92 430 Q81 430 80 420 L70 262 L52 258 L48 150 L41 232 Q40 244 32 244 Q23 244 24 232 L30 130 Q32 106 50 96 Z';

const SPOTS: Record<Exclude<SpotId, 'none' | 'body'>, { cx: number; cy: number; rx: number; ry: number }> = {
  abdomen: { cx: 110, cy: 212, rx: 52, ry: 40 },
  lower: { cx: 128, cy: 270, rx: 24, ry: 20 },
  rectum: { cx: 110, cy: 298, rx: 14, ry: 14 },
};

/** Front view of the body with the colon in position. The active symptom area is highlighted. */
export default function BodyFigure({ active }: { active: SpotId }) {
  return (
    <svg viewBox="0 0 220 440" className="body-svg" role="img" aria-label="Body diagram with colon and highlighted symptom area">
      <circle cx="110" cy="42" r="26" className="body-fill" />
      <rect x="100" y="64" width="20" height="24" className="body-fill" />
      <path d={BODY} className={`body-fill ${active === 'body' ? 'body-on' : ''}`} />
      <g transform="translate(110 235) scale(26 -26)" fill="none" strokeLinecap="round" strokeWidth={0.34} stroke="#e29bad">
        <path d="M-1.6 -1.2 L-1.62 1" />
        <path d="M-1.6 1 C-1.1 1.1 -0.6 0.5 0 0.5 S1.1 1.1 1.6 1.1" />
        <path d="M1.6 1.1 L1.6 -0.95" />
        <path d="M1.6 -0.95 C1.3 -1.4 0.9 -0.9 0.5 -1.4 S0.1 -1.6 0.05 -1.95 L0.02 -2.4" />
      </g>
      {(Object.keys(SPOTS) as (keyof typeof SPOTS)[]).map((k) => (
        <ellipse key={k} {...SPOTS[k]} className={`spot ${active === k ? 'spot-on' : ''}`} />
      ))}
    </svg>
  );
}
