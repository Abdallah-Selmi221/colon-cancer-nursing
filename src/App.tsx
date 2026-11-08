import { useEffect, useState } from 'react';
import { AnimatePresence } from 'framer-motion';
import { slides } from './content/slides';
import { usePresentation } from './engine/usePresentation';
import { PresentationContext } from './engine/context';
import SlideStage from './engine/SlideStage';
import Controls from './engine/Controls';
import Overview from './engine/Overview';
import InstitutionalHeader from './components/InstitutionalHeader';

export default function App() {
  const p = usePresentation(slides.length);
  const [idle, setIdle] = useState(false);

  // in full screen, tuck the control bar away after a pause; any pointer move brings it back
  useEffect(() => {
    if (!p.fs) { setIdle(false); return; }
    let t = window.setTimeout(() => setIdle(true), 2500);
    const wake = () => { setIdle(false); window.clearTimeout(t); t = window.setTimeout(() => setIdle(true), 2500); };
    window.addEventListener('mousemove', wake);
    return () => { window.clearTimeout(t); window.removeEventListener('mousemove', wake); };
  }, [p.fs]);

  return (
    <PresentationContext.Provider value={{ index: p.index, goTo: p.goTo, flags: p.flags }}>
      <div className={`app ${p.fs ? 'is-fs' : ''}`}>
        <InstitutionalHeader />
        <SlideStage index={p.index} flags={p.flags} />
        <Controls p={p} idle={idle} />
        <AnimatePresence>
          {p.overview && <Overview index={p.index} onPick={p.goTo} onClose={() => p.setOverview(false)} />}
        </AnimatePresence>
      </div>
    </PresentationContext.Provider>
  );
}
