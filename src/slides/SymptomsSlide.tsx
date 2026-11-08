import { useState } from 'react';
import type { Slide } from '../content/slides';
import BodyFigure from '../illustrations/BodyFigure';

export default function SymptomsSlide({ slide }: { slide: Extract<Slide, { kind: 'symptoms' }> }) {
  const [i, setI] = useState(0);
  return (
    <div className="slide-body split split-body">
      <div className="art-box body-box" data-reveal="art"><BodyFigure active={slide.items[i].spot} /></div>
      <div className="col">
        {slide.items.map((it, k) => (
          <button key={it.title} className={`sym ${k === i ? 'on' : ''}`} onClick={() => setI(k)} onMouseEnter={() => setI(k)} onFocus={() => setI(k)} data-reveal="item">
            <h2>{it.title}</h2>
            <p>{it.body}</p>
          </button>
        ))}
      </div>
    </div>
  );
}
