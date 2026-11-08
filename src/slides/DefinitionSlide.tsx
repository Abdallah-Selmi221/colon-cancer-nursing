import CrossSectionScene from '../three/CrossSectionScene';
import type { Slide } from '../content/slides';
import { Points, SceneBox } from './ui';

export default function DefinitionSlide({ slide }: { slide: Extract<Slide, { kind: 'definition' }> }) {
  return (
    <div className="slide-body split">
      <div className="col"><Points items={slide.points} /></div>
      <SceneBox><CrossSectionScene /></SceneBox>
    </div>
  );
}
