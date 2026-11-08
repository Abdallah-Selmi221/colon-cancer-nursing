import { useState } from 'react';
import type { Slide } from '../content/slides';
import { SurgeryArt, TherapyArt } from '../illustrations/TreatmentArt';
import { FileImage } from 'lucide-react';
import InfographicModal from '../components/InfographicModal';

export function SurgerySlide({ slide }: { slide: Extract<Slide, { kind: 'surgery' }> }) {
  const [modalOpen, setModalOpen] = useState(false);

  return (
    <div className="slide-body stack">
      <div className="treatment-lead-row" data-reveal="item">
        <p className="statement">{slide.lead}</p>
        <button className="infographic-pill-btn" onClick={() => setModalOpen(true)}>
          <FileImage size={16} />
          <span>View Faculty Surgical &amp; Care Poster</span>
        </button>
      </div>
      <div className="surg-grid">
        {slide.items.map((it) => (
          <article key={it.title} className="surg" data-reveal="item">
            <div className="surg-art"><SurgeryArt id={it.art} /></div>
            <h2>{it.title}</h2>
            <p>{it.body}</p>
          </article>
        ))}
      </div>
      <InfographicModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        imageSrc={`${import.meta.env.BASE_URL}images/treatment-infographic.jpg`}
        title="Management and Treatment of Colon Cancer — Surgical & Adjuvant Therapies"
      />
    </div>
  );
}

const THERAPY = ['chemo', 'radiation', 'targeted'] as const;

export function AdditionalSlide({ slide }: { slide: Extract<Slide, { kind: 'additional' }> }) {
  return (
    <div className="slide-body split split-wide">
      <div className="col">
        <p className="lead" data-reveal="item">{slide.lead}</p>
        <div className="therapies">
          {slide.items.map((t, i) => (
            <div key={t} className="therapy" data-reveal="item">
              <TherapyArt id={THERAPY[i]} /><span>{t}</span>
            </div>
          ))}
        </div>
      </div>
      <aside className="seek" data-reveal="art">
        <h2>{slide.seekTitle}</h2>
        <p>{slide.seekText}</p>
      </aside>
    </div>
  );
}

export function QuestionsSlide({ slide }: { slide: Extract<Slide, { kind: 'questions' }> }) {
  return (
    <div className="slide-body">
      <ol className="questions">
        {slide.items.map((q, i) => (
          <li key={q} data-reveal="item"><span className="q-n">{i + 1}</span>{q}</li>
        ))}
      </ol>
    </div>
  );
}
