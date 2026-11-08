import gsap from 'gsap';
import { useEffect, useRef } from 'react';
import type { Slide } from '../content/slides';

const R = 90;
const C = 2 * Math.PI * R;

export default function TypesSlide({ slide }: { slide: Extract<Slide, { kind: 'types' }> }) {
  const arc = useRef<SVGCircleElement>(null);
  const target = C * (1 - slide.percent / 100);
  useEffect(() => {
    if (!arc.current) return;
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) { arc.current.style.strokeDashoffset = String(target); return; }
    const t = gsap.fromTo(arc.current, { strokeDashoffset: C }, { strokeDashoffset: target, duration: 1.8, ease: 'power2.out', delay: 0.4 });
    return () => { t.kill(); };
  }, [target]);
  return (
    <div className="slide-body split">
      <div className="ring-wrap" data-reveal="art">
        <svg viewBox="0 0 240 240" className="ring" role="img" aria-label={`About ${slide.percent} percent`}>
          <circle cx="120" cy="120" r={R} fill="none" stroke="#f3dde3" strokeWidth="26" />
          <circle ref={arc} cx="120" cy="120" r={R} fill="none" stroke="#7a1f3d" strokeWidth="26" strokeLinecap="round"
            strokeDasharray={C} strokeDashoffset={C} transform="rotate(-90 120 120)" />
        </svg>
        <div className="ring-label"><span>about</span><strong>{slide.percent}%</strong></div>
      </div>
      <div className="col">
        <h2 className="big-term" data-reveal="item">{slide.item.title}</h2>
        <p className="lead" data-reveal="item">{slide.item.body}</p>
      </div>
    </div>
  );
}
