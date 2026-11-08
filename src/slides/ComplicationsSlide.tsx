import type { Slide } from '../content/slides';
import ComplicationArt from '../illustrations/ComplicationArt';

export default function ComplicationsSlide({ slide }: { slide: Extract<Slide, { kind: 'complications' }> }) {
  return (
    <div className="slide-body comp-grid">
      {slide.items.map((it) => (
        <article key={it.title} className="comp" data-reveal="item">
          <div className="comp-art"><ComplicationArt id={it.art} /></div>
          <h2>{it.title}</h2>
          <p>{it.body}</p>
        </article>
      ))}
    </div>
  );
}
