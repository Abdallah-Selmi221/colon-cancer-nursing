import { BookOpen, FileText } from 'lucide-react';
import type { Slide } from '../content/slides';

const ICONS: Record<string, typeof BookOpen> = {
  summary: BookOpen,
  references: FileText,
};

export default function PendingSlide({ slide }: { slide: Extract<Slide, { kind: 'pending' }> }) {
  const Icon = ICONS[slide.section] || BookOpen;
  const isSummary = slide.section === 'summary';

  return (
    <div className="slide-body center-col">
      <div className="academic-folio-card" data-reveal="art">
        <div className="folio-top-badge">
          <div className="folio-icon-circle">
            <Icon size={28} />
          </div>
          <div>
            <span className="folio-category">Faculty Academic Curriculum</span>
            <h2 className="folio-title">{slide.title}</h2>
          </div>
        </div>

        <div className="folio-divider" />

        <div className="folio-body">
          <div className="folio-pill">
            <span className="folio-dot" />
            <span>{isSummary ? 'Executive Synthesis' : 'Academic Bibliography'}</span>
          </div>

          <p className="folio-note">{slide.note}</p>

          <p className="folio-desc">
            {isSummary
              ? 'This concluding section synthesizes the core clinical concepts of colon cancer: anatomical localization, pathological progression, screening protocols, surgical/medical modalities, and evidence-based nursing care plans.'
              : 'Official clinical references, textbook citations (Brunner & Suddarth’s Medical-Surgical Nursing, NANDA-I Nursing Diagnoses, and Oncology Nursing Society Guidelines) as approved by the department faculty.'}
          </p>
        </div>

        <div className="folio-footer">
          <span className="folio-institution">كلية التمريض للبنات بالقاهرة • جامعة الأزهر</span>
          <span className="folio-tag">Medical-Surgical Nursing Department</span>
        </div>
      </div>
    </div>
  );
}
