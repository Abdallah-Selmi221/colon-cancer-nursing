import { createContext, useContext } from 'react';

export interface PresentationCtx {
  index: number;
  goTo: (i: number) => void;
  flags: boolean;
}
export const PresentationContext = createContext<PresentationCtx>({ index: 0, goTo: () => {}, flags: false });
export const usePresentationCtx = () => useContext(PresentationContext);
