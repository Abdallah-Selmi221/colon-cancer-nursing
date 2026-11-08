import { motion } from 'framer-motion';
import { FileDown, X } from 'lucide-react';
import { useEffect, useRef } from 'react';
import { slides, SECTIONS } from '../content/slides';
import { downloadPptxPresentation } from '../export/downloadPptx';

export default function Overview({ index, onPick, onClose }: { index: number; onPick: (i: number) => void; onClose: () => void }) {
  const cur = useRef<HTMLButtonElement>(null);
  useEffect(() => { cur.current?.scrollIntoView({ block: 'center' }); cur.current?.focus(); }, []);
  return (
    <motion.div className="overview" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={{ duration: 0.2 }} role="dialog" aria-label="Slide overview">
      <div className="overview-bar">
        <h2>All slides ({slides.length})</h2>
        <div className="overview-actions">
          <button
            onClick={downloadPptxPresentation}
            className="overview-export-btn"
            title="Download PowerPoint Presentation (.pptx)"
          >
            <FileDown size={17} />
            <span>Download PowerPoint (.pptx)</span>
          </button>
          <button onClick={onClose} aria-label="Close overview"><X size={20} /></button>
        </div>
      </div>
      <div className="overview-grid">
        {slides.map((s, i) => (
          <motion.button
            key={s.id}
            ref={i === index ? cur : undefined}
            className={`ov-card ${i === index ? 'current' : ''}`}
            onClick={() => { onPick(i); onClose(); }}
            whileHover={{ y: -3 }}
            transition={{ duration: 0.15 }}
          >
            <span className="ov-n">{i + 1}</span>
            <span className="ov-sec">{SECTIONS.find((x) => x.id === s.section)!.label}</span>
            <span className="ov-title">{s.title}</span>
          </motion.button>
        ))}
      </div>
    </motion.div>
  );
}
