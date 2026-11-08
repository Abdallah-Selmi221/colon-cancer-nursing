import DigestiveScene from '../three/DigestiveScene';
import type { Slide } from '../content/slides';
import { SceneBox } from './ui';
import { usePresentationCtx } from '../engine/context';
import { ChevronRight } from 'lucide-react';

export default function CoverSlide({ slide }: { slide: Extract<Slide, { kind: 'cover' }> }) {
  const { goTo } = usePresentationCtx();

  return (
    <div className="cover">
      <div className="cover-text">
        <div className="cover-badge" data-reveal="head">
          <span className="cover-badge-dot" />
          <span>Faculty of Nursing for Girls, Al-Azhar University</span>
        </div>

        <h1 className="cover-main-heading" data-reveal="head">
          <span className="cover-word-primary">COLON</span>
          <span className="cover-word-secondary">CANCER</span>
        </h1>

        <p className="cover-sub">{slide.subtitle}</p>

        <p className="cover-desc">
          A Comprehensive Clinical Seminar on Colorectal Oncology, Surgical Interventions &amp; Advanced Nursing Care Plans.
        </p>

        {/* Supervision Block */}
        <div className="cover-supervision" data-reveal="head">
          <span className="cover-supervision-label">Under Supervision</span>
          <div className="cover-supervisor-row">
            <div className="cover-supervisor">
              <span className="cover-supervisor-title">Prof. Dr. Heba Abdel-Azem</span>
              <span className="cover-supervisor-role">Professor of Medical Surgical Nursing Department</span>
            </div>
            <div className="cover-supervisor-divider" />
            <div className="cover-supervisor">
              <span className="cover-supervisor-title">Dr. Eman Abel-Maqsoud</span>
              <span className="cover-supervisor-role">Demonstrator of Medical Surgical Nursing Department</span>
            </div>
          </div>
          <span className="cover-year">Academic Year 2026 / 2027</span>
        </div>

        <div className="cover-cta-row">
          <button className="cover-start-btn" onClick={() => goTo(1)}>
            <span>Begin Presentation</span>
            <ChevronRight size={18} />
          </button>
          <span className="cover-hint">Press <kbd>→</kbd> or <kbd>Space</kbd></span>
        </div>
      </div>

      <SceneBox className="cover-art" data-reveal="art"><DigestiveScene /></SceneBox>
    </div>
  );
}
