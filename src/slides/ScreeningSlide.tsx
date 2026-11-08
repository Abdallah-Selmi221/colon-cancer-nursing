import type { Slide } from '../content/slides';
import { Points } from './ui';

export default function ScreeningSlide({ slide }: { slide: Extract<Slide, { kind: 'screening' }> }) {
  return (
    <div className="slide-body split">
      <div className="col"><Points items={slide.points} /></div>
      <div className="art-box timeline-box" data-reveal="art">
        <svg viewBox="0 0 440 220" className="timeline" role="img" aria-label="Polyps take 10 to 15 years to turn into cancer; screening allows removal before they develop">
          <path d="M40 130 H400" stroke="#e8c6cf" strokeWidth="8" strokeLinecap="round" fill="none" />
          <path d="M40 130 H400" stroke="#7a1f3d" strokeWidth="8" strokeLinecap="round" fill="none" data-draw />
          <circle cx="40" cy="130" r="13" fill="#f2b6c4" stroke="#c47e92" strokeWidth="3" />
          <circle cx="400" cy="130" r="17" fill="#7a1f3d" />
          <text x="40" y="170" textAnchor="middle" className="tl-t">Polyps</text>
          <text x="400" y="170" textAnchor="middle" className="tl-t">Cancer</text>
          <path d="M40 98 H400" stroke="#a56f80" strokeWidth="2" />
          <path d="M40 90 V106 M400 90 V106" stroke="#a56f80" strokeWidth="2" />
          <text x="220" y="84" textAnchor="middle" className="tl-b">10 to 15 years</text>
          <path d="M150 130 V196" stroke="#c9567b" strokeWidth="3" strokeDasharray="5 4" />
          <circle cx="150" cy="130" r="9" fill="#fff" stroke="#c9567b" strokeWidth="4" />
          <text x="156" y="212" className="tl-t">Screening allows removal</text>
        </svg>
      </div>
    </div>
  );
}
