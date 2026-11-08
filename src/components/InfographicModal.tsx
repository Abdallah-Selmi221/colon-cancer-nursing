import { X, ZoomIn, Maximize2 } from 'lucide-react';
import { useEffect, useState } from 'react';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  imageSrc: string;
  title: string;
}

export default function InfographicModal({ isOpen, onClose, imageSrc, title }: Props) {
  const [zoomed, setZoomed] = useState(false);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  return (
    <div className="infographic-overlay" onClick={onClose} role="dialog" aria-modal="true" aria-label={title}>
      <div className={`infographic-dialog ${zoomed ? 'is-zoomed' : ''}`} onClick={(e) => e.stopPropagation()}>
        <header className="infographic-head">
          <div className="infographic-head-title">
            <span className="infographic-badge">Faculty Clinical Poster</span>
            <h3>{title}</h3>
          </div>
          <div className="infographic-actions">
            <button
              className="infographic-btn"
              onClick={() => setZoomed((z) => !z)}
              title={zoomed ? 'Original Size' : 'Zoom In'}
              aria-label="Toggle zoom"
            >
              <ZoomIn size={18} />
            </button>
            <button className="infographic-btn close-btn" onClick={onClose} title="Close (Esc)" aria-label="Close modal">
              <X size={20} />
            </button>
          </div>
        </header>

        <div className="infographic-viewport">
          <img src={imageSrc} alt={title} className={`infographic-img ${zoomed ? 'zoomed' : ''}`} />
        </div>
      </div>
    </div>
  );
}
