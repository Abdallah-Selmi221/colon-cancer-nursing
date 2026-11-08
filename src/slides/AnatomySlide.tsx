import { useState } from 'react';
import { ANATOMY_SECTIONS, type SectionKey, type Slide } from '../content/slides';
import { Points, SceneBox } from './ui';
import AnatomyScene from '../three/AnatomyScene';

export default function AnatomySlide({ slide }: { slide: Extract<Slide, { kind: 'anatomy' }> }) {
  const [active, setActive] = useState<SectionKey>('ascending');
  if (slide.mode === 'facts') {
    return (
      <div className="slide-body split">
        <div className="col">
          <Points items={slide.facts ?? []} />
          {slide.statement && <p className="statement" data-reveal="item">{slide.statement}</p>}
        </div>
        <SceneBox><AnatomyScene /></SceneBox>
      </div>
    );
  }
  const cur = ANATOMY_SECTIONS.find((s) => s.key === active)!;
  return (
    <div className="slide-body split split-rev">
      <SceneBox><AnatomyScene interactive active={active} onSelect={setActive} /></SceneBox>
      <div className="col">
        <div className="seg-list" role="tablist" aria-label="Colon sections">
          {ANATOMY_SECTIONS.map((s) => (
            <button key={s.key} role="tab" aria-selected={s.key === active} className={`seg ${s.key === active ? 'on' : ''}`} onClick={() => setActive(s.key)} data-reveal="item">
              {s.label}
            </button>
          ))}
        </div>
        <div className="seg-detail" data-reveal="item" key={cur.key}>
          <h2>{cur.label}</h2>
          <p>{cur.text}</p>
        </div>
        <p className="hint" data-reveal="item">Select a section here or on the model. Drag the model to turn it.</p>
      </div>
    </div>
  );
}
