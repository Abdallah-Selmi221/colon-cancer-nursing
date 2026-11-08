import { useState, type CSSProperties } from 'react';
import { STAGE_RAIL, type Slide } from '../content/slides';
import StageArt from '../illustrations/StageArt';
import { FileImage } from 'lucide-react';
import InfographicModal from '../components/InfographicModal';

export default function StagesSlide({ slide }: { slide: Extract<Slide, { kind: 'stages' }> }) {
  const [modalOpen, setModalOpen] = useState(false);
  const here = slide.items.map((i) => i.stage);

  return (
    <div className="slide-body stack">
      <div className="stages-header-row" data-reveal="item">
        <ol className="rail" aria-label="Stage progression">
          {STAGE_RAIL.map((s) => (
            <li key={s} className={here.includes(s) ? 'on' : ''}>
              {s}
            </li>
          ))}
        </ol>
        <button className="infographic-pill-btn" onClick={() => setModalOpen(true)}>
          <FileImage size={16} />
          <span>Illustrated Staging Board</span>
        </button>
      </div>

      <div className="stage-cards" style={{ '--n': slide.items.length } as CSSProperties}>
        {slide.items.map((it) => (
          <article key={it.stage} className="stage-card" data-reveal="item">
            <div className="stage-art"><StageArt stage={it.stage} /></div>
            <h2>Stage {it.stage}</h2>
            <p>{it.text}</p>
          </article>
        ))}
      </div>

      <InfographicModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        imageSrc={`${import.meta.env.BASE_URL}images/diag-stages-infographic.jpg`}
        title="Colorectal Cancer Staging — Visual Anatomy & Distant Spread"
      />
    </div>
  );
}
