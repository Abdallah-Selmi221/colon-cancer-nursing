import type { Slide } from '../content/slides';
import { AbdomenFigure } from '../illustrations/DiagnosticArt';
import { StomaArt } from '../illustrations/TreatmentArt';

export default function StepsSlide({ slide }: { slide: Extract<Slide, { kind: 'steps' }> }) {
  const list = (
    <div className="col">
      <p className="label" data-reveal="item">{slide.phase}</p>
      {slide.items.map((it) => (
        <article key={it.title} className="step-card" data-reveal="item">
          <h2>{it.title}</h2>
          {it.body && <p>{it.body}</p>}
        </article>
      ))}
    </div>
  );
  if (!slide.art) return <div className="slide-body steps-wide">{list}</div>;
  return (
    <div className="slide-body split split-wide">
      {list}
      <div className="art-box exam-box" data-reveal="art"><AbdomenFigure><StomaArt mode={slide.art === 'site' ? 'site' : 'stoma'} /></AbdomenFigure></div>
    </div>
  );
}
