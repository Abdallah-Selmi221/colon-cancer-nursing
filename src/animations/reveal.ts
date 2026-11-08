import gsap from 'gsap';
import { useEffect, useState } from 'react';

export function usePrefersReducedMotion() {
  const [reduced, setReduced] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(prefers-reduced-motion: reduce)').matches,
  );
  useEffect(() => {
    const mq = window.matchMedia('(prefers-reduced-motion: reduce)');
    const on = () => setReduced(mq.matches);
    mq.addEventListener('change', on);
    return () => mq.removeEventListener('change', on);
  }, []);
  return reduced;
}

/** Snappy, elegant entrance choreography for every slide. Elements opt in with data-reveal. */
export function revealSlide(root: HTMLElement) {
  const tl = gsap.timeline({ defaults: { ease: 'power2.out' } });

  const add = (sel: string, fromVars: gsap.TweenVars, toVars: gsap.TweenVars, at: number | string) => {
    const els = root.querySelectorAll(sel);
    if (els.length) tl.fromTo(els, fromVars, { ...toVars, clearProps: 'opacity,transform' }, at);
  };

  add('[data-reveal="head"]', { y: 10, opacity: 0 }, { y: 0, opacity: 1, duration: 0.28, stagger: 0.03 }, 0);
  add('[data-reveal="art"]', { scale: 0.97, opacity: 0 }, { scale: 1, opacity: 1, duration: 0.35 }, 0.02);
  add('[data-reveal="cards"]', { y: 12, opacity: 0 }, { y: 0, opacity: 1, duration: 0.3 }, 0.03);
  add('[data-reveal="item"]', { y: 12, opacity: 0 }, { y: 0, opacity: 1, duration: 0.25, stagger: 0.02 }, 0.04);
  add('[data-reveal="line"]', { scaleX: 0, transformOrigin: 'left center' }, { scaleX: 1, duration: 0.35, stagger: 0.04 }, 0.05);

  root.querySelectorAll<SVGGeometryElement>('[data-draw]').forEach((el, i) => {
    const len = el.getTotalLength();
    el.style.strokeDasharray = String(len);
    tl.fromTo(el, { strokeDashoffset: len }, { strokeDashoffset: 0, duration: 0.6, ease: 'power1.inOut' }, 0.1 + i * 0.04);
  });

  return tl;
}
