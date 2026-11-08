import DigestiveScene from '../three/DigestiveScene';
import type { Slide } from '../content/slides';
import { Chip, Points, SceneBox } from './ui';

export default function IntroSlide({ slide }: { slide: Extract<Slide, { kind: 'intro' }> }) {
  return (
    <div className={`slide-body ${slide.visual ? 'split' : 'center-col'}`}>
      <div className="col">
        {slide.points && <Points items={slide.points} />}
        {slide.chips && (
          <div data-reveal="item">
            <p className="label">{slide.chipsLabel}</p>
            <div className="chips">{slide.chips.map((c) => <Chip key={c}>{c}</Chip>)}</div>
          </div>
        )}
        {slide.statement && <p className="statement" data-reveal="item">{slide.statement}</p>}
      </div>
      {slide.visual && <SceneBox><DigestiveScene /></SceneBox>}
    </div>
  );
}
