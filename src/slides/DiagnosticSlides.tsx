import { useState } from 'react';
import { DIAGNOSTIC_TESTS, type Slide } from '../content/slides';
import { BloodArt, ColonoscopyArt, ExamArt, ImagingArt } from '../illustrations/DiagnosticArt';
import { Chip } from './ui';
import { FileImage } from 'lucide-react';
import InfographicModal from '../components/InfographicModal';

export function ExamSlide({ slide }: { slide: Extract<Slide, { kind: 'exam' }> }) {
  return (
    <div className="slide-body split">
      <div className="col">
        <ol className="flow">
          {slide.steps.map((s) => <li key={s} data-reveal="item"><p>{s}</p></li>)}
        </ol>
      </div>
      <div className="art-box exam-box" data-reveal="art"><ExamArt /></div>
    </div>
  );
}

const ART = { blood: BloodArt, imaging: ImagingArt, colonoscopy: ColonoscopyArt };

export function TestsSlide() {
  const [modalOpen, setModalOpen] = useState(false);

  return (
    <div className="slide-body">
      <div className="subhead-action" data-reveal="item">
        <button className="infographic-pill-btn" onClick={() => setModalOpen(true)}>
          <FileImage size={16} />
          <span>View Faculty Diagnosis &amp; Staging Poster</span>
        </button>
      </div>
      <div className="tests-grid">
        {DIAGNOSTIC_TESTS.map((t) => {
          const Art = ART[t.id];
          return (
            <article key={t.id} className="test" data-reveal="item">
              <div className="test-art"><Art /></div>
              <h2>{t.title}</h2>
              <p>{t.text}</p>
              {t.chips.length > 0 && <div className="chips">{t.chips.map((c) => <Chip key={c}>{c}</Chip>)}</div>}
            </article>
          );
        })}
      </div>
      <InfographicModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        imageSrc={`${import.meta.env.BASE_URL}images/diag-stages-infographic.jpg`}
        title="Diagnosis and Tests & Cancer Stages — Faculty Reference"
      />
    </div>
  );
}

export function WhyTestsSlide({ slide }: { slide: Extract<Slide, { kind: 'whyTests' }> }) {
  return (
    <div className="slide-body center-col">
      <p className="statement" data-reveal="item">{slide.text}</p>
      <div className="why" data-reveal="item">
        <div className="why-from">
          {DIAGNOSTIC_TESTS.map((t) => <span key={t.id} className="why-pill">{t.title}</span>)}
        </div>
        <svg viewBox="0 0 80 120" className="why-arrow" aria-hidden><path d="M6 20 H34 Q44 20 44 34 V50 M6 100 H34 Q44 100 44 86 V70 M6 60 H70" fill="none" stroke="#a56f80" strokeWidth="3" data-draw /></svg>
        <div className="why-to"><span className="why-out">Diagnose cancer</span><span className="why-out">Determine its stage</span></div>
      </div>
    </div>
  );
}
