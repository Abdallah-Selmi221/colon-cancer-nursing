import gsap from 'gsap';
import { useCallback, useEffect, useRef, useState } from 'react';
import { PATHO_STAGES, type Slide } from '../content/slides';
import { Chip, SceneBox } from './ui';
import { Play } from 'lucide-react';
import PathoScene from '../three/PathoScene';

function Genes({ slide }: { slide: Extract<Slide, { kind: 'patho' }> }) {
  return (
    <div className="slide-body center-col">
      <ol className="flow">
        {(slide.points ?? []).map((p, i) => (
          <li key={p} data-reveal="item">
            <p>{p}</p>
            {i === 0 && <div className="chips">{slide.genes?.map((g) => <Chip key={g}>{g}</Chip>)}</div>}
          </li>
        ))}
      </ol>
    </div>
  );
}

function Progression() {
  const [stage, setStage] = useState(0);
  const tl = useRef<gsap.core.Timeline | null>(null);
  const stop = () => { tl.current?.kill(); tl.current = null; };
  const play = useCallback(() => {
    tl.current?.kill();
    setStage(0);
    const t = gsap.timeline();
    [1, 2, 3].forEach((s, i) => t.call(() => setStage(s), [], 1.6 + i * 2.4));
    tl.current = t;
  }, []);
  useEffect(() => {
    if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches) play();
    return () => { tl.current?.kill(); };
  }, [play]);
  return (
    <div className="slide-body split split-wide">
      <SceneBox><PathoScene stage={stage} /></SceneBox>
      <div className="col">
        <ol className="stage-steps">
          {PATHO_STAGES.map((s, i) => (
            <li key={s} data-reveal="item">
              <button className={`step ${i === stage ? 'on' : i < stage ? 'done' : ''}`} onClick={() => { stop(); setStage(i); }}>
                <span className="step-n">{i + 1}</span>{s}
              </button>
            </li>
          ))}
        </ol>
        <button className="ghost-btn" onClick={play} data-reveal="item"><Play size={16} /> Play progression</button>
      </div>
    </div>
  );
}

export default function PathoSlide({ slide }: { slide: Extract<Slide, { kind: 'patho' }> }) {
  return slide.mode === 'genes' ? <Genes slide={slide} /> : <Progression />;
}
