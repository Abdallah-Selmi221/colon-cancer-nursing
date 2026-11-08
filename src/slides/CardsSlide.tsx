import type { CSSProperties } from 'react';
import type { Slide } from '../content/slides';

export default function CardsSlide({ slide }: { slide: Extract<Slide, { kind: 'cards' }> }) {
  if (slide.layout === 'rows') {
    return (
      <div className="slide-body">
        <ul className="rows">
          {slide.items.map((it) => (
            <li key={it.title} data-reveal="item"><span className="row-bar" />{it.title}</li>
          ))}
        </ul>
      </div>
    );
  }
  return (
    <div className="slide-body">
      <div className="grid" style={{ '--cols': slide.columns ?? 4 } as CSSProperties}>
        {slide.items.map((it) => (
          <article key={it.title} className="card" data-reveal="item">
            <h2>{it.title}</h2>
            {it.body && <p>{it.body}</p>}
          </article>
        ))}
      </div>
    </div>
  );
}
