import { SECTIONS, firstIndexOf } from '../content/slides';
import { usePresentationCtx } from '../engine/context';

export default function OutlineSlide() {
  const { goTo, index } = usePresentationCtx();

  return (
    <div className="slide-body outline-slide-body">
      <ol className="outline" data-reveal="cards">
        {SECTIONS.map((s, i) => {
          const slideIdx = firstIndexOf(s.id);
          const isCurrent = slideIdx === index;
          return (
            <li key={s.id} data-reveal="item">
              <button
                type="button"
                onClick={() => goTo(slideIdx)}
                aria-current={isCurrent ? 'true' : undefined}
                className="outline-btn"
              >
                <span className="outline-n">{i + 1}</span>
                <span className="outline-label">{s.label}</span>
              </button>
            </li>
          );
        })}
      </ol>
    </div>
  );
}
