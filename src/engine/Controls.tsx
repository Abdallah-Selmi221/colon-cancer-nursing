import { ChevronLeft, ChevronRight, Flag, LayoutGrid, Maximize2, Minimize2 } from 'lucide-react';
import { slides, SECTIONS } from '../content/slides';
import type { Presentation } from './usePresentation';

export default function Controls({ p, idle }: { p: Presentation; idle: boolean }) {
  const slide = slides[p.index];
  const next = slides[p.index + 1];
  const sec = SECTIONS.find((s) => s.id === slide.section)!;
  const pct = ((p.index + 1) / p.total) * 100;
  return (
    <footer className={`controls ${idle ? 'idle' : ''}`}>
      <div className="where">
        <span className="where-sec">{sec.label}</span>
        <span className="where-next">{next ? `Next: ${next.title}` : 'End of presentation'}</span>
      </div>
      <div className="progress" role="progressbar" aria-valuemin={1} aria-valuemax={p.total} aria-valuenow={p.index + 1} aria-label="Slide progress">
        <span style={{ width: `${pct}%` }} />
      </div>
      <div className="ctl-btns">
        <span className="counter">{p.index + 1} / {p.total}</span>
        <button onClick={() => p.setFlags((f) => !f)} aria-pressed={p.flags} title="Show review flags (R)" className={p.flags ? 'on' : ''}><Flag size={18} /></button>
        <button onClick={() => p.setOverview(true)} title="Slide overview (O)" aria-label="Slide overview"><LayoutGrid size={18} /></button>
        <button onClick={p.toggleFs} title="Full screen (F)" aria-label="Toggle full screen">{p.fs ? <Minimize2 size={18} /> : <Maximize2 size={18} />}</button>
        <button onClick={p.prev} disabled={p.index === 0} title="Previous (←)" aria-label="Previous slide"><ChevronLeft size={20} /></button>
        <button onClick={p.next} disabled={p.index === p.total - 1} title="Next (→ or Space)" aria-label="Next slide" className="primary"><ChevronRight size={20} /></button>
      </div>
    </footer>
  );
}
