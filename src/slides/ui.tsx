import { ReactNode } from 'react';

export const Chip = ({ children }: { children: ReactNode }) => <span className="chip">{children}</span>;

export function Points({ items }: { items: string[] }) {
  return (
    <ul className="points">
      {items.map((t) => (
        <li key={t} data-reveal="item"><span className="dot" />{t}</li>
      ))}
    </ul>
  );
}

/** Box that hosts an interactive 3D scene. */
export function SceneBox({ children, className = '' }: { children: ReactNode; className?: string }) {
  return (
    <div className={`art-box ${className}`} data-reveal="art">
      {children}
    </div>
  );
}
