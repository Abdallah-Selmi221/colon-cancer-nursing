import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import gsap from 'gsap';
import { slides, SECTIONS } from '../content/slides';
import { REGISTRY } from './registry';
import { revealSlide, usePrefersReducedMotion } from '../animations/reveal';

export default function SlideStage({ index, flags }: { index: number; flags: boolean }) {
  const [shown, setShown] = useState(index);
  const ref = useRef<HTMLDivElement>(null);
  const reduced = usePrefersReducedMotion();

  // exit: fade the current slide out, then swap
  useEffect(() => {
    if (index === shown) return;
    const el = ref.current;
    if (!el || reduced) { setShown(index); return; }
    const dir = index > shown ? 1 : -1;
    const tween = gsap.to(el, { opacity: 0, x: -36 * dir, duration: 0.26, ease: 'power2.in', onComplete: () => setShown(index) });
    return () => { tween.kill(); };
  }, [index, shown, reduced]);

  // enter: one choreographed timeline per slide
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    gsap.set(el, { opacity: 1, x: 0 });
    if (reduced) return;
    const ctx = gsap.context(() => { revealSlide(el); }, el);
    return () => ctx.revert();
  }, [shown, reduced]);

  const slide = slides[shown];
  const Comp = REGISTRY[slide.kind];
  const secNo = SECTIONS.findIndex((s) => s.id === slide.section);
  const sec = SECTIONS[secNo];

  return (
    <main className="stage" aria-live="polite">
      <div ref={ref} className={`stage-inner kind-${slide.kind}`} key={slide.id}>
        {slide.kind !== 'cover' && (
          <header className="slide-head">
            <p className="eyebrow" data-reveal="head">{sec.label}</p>
            <h1 data-reveal="head">{slide.title}</h1>
          </header>
        )}
        <Comp slide={slide} />
        {flags && slide.flag && (
          <aside className="flag-ribbon" role="note"><strong>Review:</strong> {slide.flag}</aside>
        )}
      </div>
    </main>
  );
}
