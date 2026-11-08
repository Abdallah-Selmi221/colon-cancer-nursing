import { useCallback, useEffect, useState } from 'react';

const parse = (total: number) => {
  const m = window.location.hash.match(/#\/(\d+)/);
  const n = m ? parseInt(m[1], 10) - 1 : 0;
  return Math.min(Math.max(n, 0), total - 1);
};

export function usePresentation(total: number) {
  const [index, setIndex] = useState(() => parse(total));
  const [overview, setOverview] = useState(false);
  const [fs, setFs] = useState(false);
  const [flags, setFlags] = useState(false);

  const next = useCallback(() => setIndex((i) => Math.min(total - 1, i + 1)), [total]);
  const prev = useCallback(() => setIndex((i) => Math.max(0, i - 1)), []);
  const goTo = useCallback((n: number) => setIndex(Math.min(total - 1, Math.max(0, n))), [total]);
  const toggleFs = useCallback(() => {
    if (document.fullscreenElement) void document.exitFullscreen();
    else void document.documentElement.requestFullscreen?.()?.catch(() => {});
  }, []);

  useEffect(() => {
    history.replaceState(null, '', `#/${index + 1}`);
  }, [index]);

  useEffect(() => {
    const onHash = () => setIndex(parse(total));
    const onFs = () => setFs(!!document.fullscreenElement);
    window.addEventListener('hashchange', onHash);
    document.addEventListener('fullscreenchange', onFs);
    return () => {
      window.removeEventListener('hashchange', onHash);
      document.removeEventListener('fullscreenchange', onFs);
    };
  }, [total]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      const t = e.target as HTMLElement | null;
      if (e.key === ' ' && t?.closest('button,a,input,select,textarea')) return;
      switch (e.key) {
        case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ':
          e.preventDefault(); if (!overview) next(); break;
        case 'ArrowLeft': case 'ArrowUp': case 'PageUp':
          e.preventDefault(); if (!overview) prev(); break;
        case 'Home': goTo(0); break;
        case 'End': goTo(total - 1); break;
        case 'o': case 'O': setOverview((o) => !o); break;
        case 'f': case 'F': toggleFs(); break;
        case 'r': case 'R': setFlags((f) => !f); break;
        case 'Escape': setOverview(false); break;
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [overview, next, prev, goTo, total, toggleFs]);

  return { index, next, prev, goTo, overview, setOverview, fs, toggleFs, flags, setFlags, total };
}
export type Presentation = ReturnType<typeof usePresentation>;
