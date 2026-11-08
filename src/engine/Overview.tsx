import { motion } from 'framer-motion';
import { X } from 'lucide-react';
import { useEffect, useRef } from 'react';
import { slides, SECTIONS } from '../content/slides';

export default function Overview({ index, onPick, onClose }: { index: number; onPick: (i: number) => void; onClose: () => void }) {
  const cur = useRef<HTMLButtonElement>(null);
  useEffect(() => { cur.current?.scrollIntoView({ block: 'center' }); cur.current?.focus(); }, []);
  return (
    <motion.div className="overview" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} transition={{ duration: 0.2 }} role="dialog" aria-label="Slide overview">
      <div className="overview-bar">
        <h2>All slides</h2>
        <button onClick={onClose} aria-label="Close overview"><X size={20} /></button>
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
