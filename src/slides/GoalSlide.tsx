import { Target, ShieldCheck, HeartPulse, Sparkles, BookOpen } from 'lucide-react';
import type { Slide } from '../content/slides';

export default function GoalSlide({ slide }: { slide: Extract<Slide, { kind: 'goal' }> }) {
  const domains = [
    'Anatomy & Physiology',
    'Pathophysiology',
    'Risk Factors & Causes',
    'Clinical Manifestations',
    'Diagnosis & Staging',
    'Management & Treatment',
    'Evidence-Based Nursing Care',
  ];

  return (
    <div className="slide-body goal-slide-body">
      <div className="goal-layout" data-reveal="cards">
        {/* Main Academic Statement Card */}
        <div className="goal-card-main">
          <div className="goal-pill">
            <Target size={18} />
            <span>PRIMARY ASSIGNMENT GOAL</span>
          </div>

          <blockquote className="goal-statement">
            "{slide.text}"
          </blockquote>

          <div className="goal-domains-section">
            <span className="goal-domains-label">Key Competency Domains Covered:</span>
            <div className="goal-domains-chips">
              {domains.map((dom) => (
                <span key={dom} className="goal-chip">
                  <ShieldCheck size={14} className="goal-chip-icon" />
                  {dom}
                </span>
              ))}
            </div>
          </div>
        </div>

        {/* Visual Clinical Panel */}
        <div className="goal-visual-panel" data-reveal="art">
          <div className="goal-visual-card">
            <div className="goal-icon-glow">
              <HeartPulse size={48} className="goal-pulse-icon" />
            </div>
            <h3 className="goal-visual-title">Clinical Nursing Excellence</h3>
            <p className="goal-visual-desc">
              Safe, holistic, and evidence-guided nursing assessment and interventions for colorectal cancer patients.
            </p>
            <div className="goal-faculty-cite">
              <BookOpen size={16} />
              <span>Department of Medical-Surgical Nursing</span>
            </div>
            <div className="goal-sparkle-row">
              <Sparkles size={16} />
              <span>Faculty of Nursing for Girls • Al-Azhar University</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
