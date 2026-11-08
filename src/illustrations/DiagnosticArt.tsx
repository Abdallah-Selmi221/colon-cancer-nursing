import gsap from 'gsap';
import { ReactNode, useEffect, useRef } from 'react';

const WALL = '#c47e92', INK = '#8d6a77';

/** Torso (front) used by the exam, stoma-site and stoma illustrations. */
export function AbdomenFigure({ children }: { children?: ReactNode }) {
  return (
    <svg viewBox="0 0 260 230" className="abdomen-svg" role="img" aria-label="Abdomen">
      <path d="M60 20 Q130 6 200 20 L215 110 Q210 170 190 215 L70 215 Q50 170 45 110 Z" fill="#fbe4e8" stroke="#d99aac" strokeWidth="3" />
      <circle cx="130" cy="118" r="4" fill="#d99aac" />
      {children}
    </svg>
  );
}

export function ExamArt() {
  return (
    <AbdomenFigure>
      <path d="M130 30 V205 M58 118 H202" stroke="#d99aac" strokeWidth="1.5" strokeDasharray="5 5" />
      <g transform="rotate(-18 168 150)" fill="#d7b5c4" stroke="#a56f80" strokeWidth="2">
        <ellipse cx="168" cy="150" rx="36" ry="21" />
        {[-14, -5, 4, 13].map((dy) => <ellipse key={dy} cx="142" cy={150 + dy} rx="20" ry="5.5" />)}
      </g>
    </AbdomenFigure>
  );
}

export function BloodArt() {
  const fill = useRef<SVGRectElement>(null);
  useEffect(() => {
    if (!fill.current || window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const t = gsap.fromTo(fill.current, { attr: { y: 176, height: 0 } }, { attr: { y: 96, height: 80 }, duration: 1.6, ease: 'power2.out', delay: 0.5 });
    return () => { t.kill(); };
  }, []);
  return (
    <svg viewBox="0 0 160 200" className="test-svg" role="img" aria-label="Blood sample tube">
      <defs><clipPath id="tube"><path d="M62 40 H98 V160 A18 18 0 0 1 62 160 Z" /></clipPath></defs>
      <rect x="58" y="22" width="44" height="22" rx="5" fill="#a56f80" />
      <path d="M62 40 H98 V160 A18 18 0 0 1 62 160 Z" fill="#fff" stroke={WALL} strokeWidth="3" />
      <g clipPath="url(#tube)"><rect ref={fill} x="60" y="96" width="40" height="80" fill="#9d1f3f" /></g>
      <path d="M68 50 V150" stroke="#fff" strokeWidth="3" opacity="0.6" strokeLinecap="round" />
    </svg>
  );
}

export function ImagingArt() {
  return (
    <svg viewBox="0 0 220 190" className="test-svg" role="img" aria-label="CT scanner">
      <rect x="40" y="14" width="140" height="130" rx="30" fill="#f1dfe4" stroke="#d99aac" strokeWidth="3" />
      <circle cx="110" cy="80" r="46" fill="#fff" stroke="#c9a2b0" strokeWidth="3" />
      <ellipse cx="110" cy="84" rx="34" ry="26" fill="#f6d4db" stroke="#d99aac" strokeWidth="2" />
      <circle cx="110" cy="96" r="4" fill="#b88c9b" />
      {[[92, 76], [108, 70], [126, 76], [124, 90], [96, 88]].map(([x, y]) => <circle key={`${x}${y}`} cx={x} cy={y} r="6" fill="none" stroke="#c47e92" strokeWidth="2.5" />)}
      <rect x="20" y="146" width="180" height="14" rx="7" fill="#d9b8c4" />
      <rect x="86" y="160" width="48" height="22" rx="4" fill="#c9a2b0" />
    </svg>
  );
}

export function ColonoscopyArt() {
  const g = useRef<SVGGElement>(null);
  useEffect(() => {
    if (!g.current || window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const t = gsap.fromTo(g.current, { scale: 0.86, svgOrigin: '110 86' }, { scale: 1.12, svgOrigin: '110 86', duration: 4, ease: 'power1.inOut', delay: 0.4 });
    return () => { t.kill(); };
  }, []);
  return (
    <svg viewBox="0 0 220 190" className="test-svg" role="img" aria-label="Colonoscopy view of the colon lumen">
      <defs>
        <radialGradient id="lumen" cx="50%" cy="50%" r="50%">
          <stop offset="0" stopColor="#3a0a1a" /><stop offset="0.35" stopColor="#8a2a42" /><stop offset="1" stopColor="#e08a96" />
        </radialGradient>
        <clipPath id="view"><circle cx="110" cy="86" r="66" /></clipPath>
      </defs>
      <rect x="22" y="8" width="176" height="154" rx="14" fill="#f1dfe4" stroke="#d99aac" strokeWidth="3" />
      <circle cx="110" cy="86" r="70" fill="#fff" stroke="#c9a2b0" strokeWidth="3" />
      <g clipPath="url(#view)">
        <g ref={g}>
          <circle cx="110" cy="86" r="90" fill="url(#lumen)" />
          {[78, 60, 44, 31, 20, 11].map((r, i) => (
            <ellipse key={r} cx={110 - i * 1.2} cy={86 + i * 0.8} rx={r} ry={r * 0.96} fill="none" stroke="#f3b3bd" strokeOpacity={0.55 - i * 0.05} strokeWidth={3.5 - i * 0.35} />
          ))}
          <circle cx="82" cy="104" r="14" fill="#d4627a" /><circle cx="78" cy="100" r="5" fill="#f4b6c0" opacity="0.8" />
        </g>
      </g>
      <path d="M198 150 C214 160 210 180 186 182 H120" fill="none" stroke={INK} strokeWidth="6" strokeLinecap="round" />
    </svg>
  );
}
