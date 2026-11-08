import { useState } from 'react';
import type { Slide } from '../content/slides';
import { Points, SceneBox } from './ui';
import { FileImage } from 'lucide-react';
import InfographicModal from '../components/InfographicModal';
import AnatomyScene from '../three/AnatomyScene';

export default function PrognosisSlide({ slide }: { slide: Extract<Slide, { kind: 'prognosis' }> }) {
  const [modalOpen, setModalOpen] = useState(false);

  if (slide.mode === 'outlook') {
    return (
      <div className="slide-body">
        <div className="subhead-action" data-reveal="item">
          <button className="infographic-pill-btn" onClick={() => setModalOpen(true)}>
            <FileImage size={16} />
            <span>View Faculty Outlook &amp; Prognosis Poster</span>
          </button>
        </div>
        <div className="split">
          <div className="col"><Points items={slide.points} /></div>
          <SceneBox><AnatomyScene tumor /></SceneBox>
        </div>
        <InfographicModal
          isOpen={modalOpen}
          onClose={() => setModalOpen(false)}
          imageSrc={`${import.meta.env.BASE_URL}images/prognosis-infographic.jpg`}
          title="Outlook & Prognosis of Colon Cancer — Early Detection & Recovery"
        />
      </div>
    );
  }
  return (
    <div className="slide-body stack">
      <div className="stages-header-row" data-reveal="item">
        <span className="label">Clinical Trajectory</span>
        <button className="infographic-pill-btn" onClick={() => setModalOpen(true)}>
          <FileImage size={16} />
          <span>Faculty Prognosis Reference Card</span>
        </button>
      </div>
      <div className="split split-tight">
        <div className="col"><Points items={slide.points} /></div>
        <div className="col">
          <div className="gauge" data-reveal="art">
            <div className="gauge-bar" />
            <div className="gauge-labels"><span>Early-stage cancer<br /><strong>Better prognosis</strong></span><span>Advanced cancer<br /><strong>Poorer prognosis</strong></span></div>
          </div>
        </div>
      </div>
      {slide.takeaway && (
        <aside className="takeaway" data-reveal="item"><h2>{slide.takeawayLabel}</h2><p>{slide.takeaway}</p></aside>
      )}
      <InfographicModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        imageSrc={`${import.meta.env.BASE_URL}images/prognosis-infographic.jpg`}
        title="Outlook & Prognosis of Colon Cancer — Early Detection & Recovery"
      />
    </div>
  );
}
