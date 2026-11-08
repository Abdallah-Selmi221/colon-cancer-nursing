import { CheckCircle2, BookmarkCheck } from 'lucide-react';
import type { Slide } from '../content/slides';

export default function ObjectivesSlide({ slide }: { slide: Extract<Slide, { kind: 'objectives' }> }) {
  return (
    <div className="slide-body objectives-slide-body">
      <div className="obj-lead-bar" data-reveal="head">
        <div className="obj-lead-pill">
          <BookmarkCheck size={18} />
          <span>LEARNING OUTCOMES</span>
        </div>
        <p className="obj-lead-text">{slide.lead}</p>
      </div>

      <div className="obj-list" data-reveal="cards">
        {slide.items.map((item, idx) => (
          <div key={idx} className="obj-card">
            <div className="obj-num-badge">
              <span>{idx + 1}</span>
            </div>
            <div className="obj-card-content">
              <p className="obj-card-text">{item}</p>
            </div>
            <CheckCircle2 size={18} className="obj-check-icon" />
          </div>
        ))}
      </div>
    </div>
  );
}
