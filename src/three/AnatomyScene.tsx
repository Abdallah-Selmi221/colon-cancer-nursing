import { OrbitControls } from '@react-three/drei';
import type { SectionKey } from '../content/slides';
import ColonFallback from '../illustrations/ColonFallback';
import ColonModel from './ColonModel';
import SceneFrame from './SceneFrame';

interface Props { active?: SectionKey | null; onSelect?: (k: SectionKey) => void; interactive?: boolean; tumor?: boolean }

export default function AnatomyScene({ active = null, onSelect, interactive = false, tumor = false }: Props) {
  return (
    <SceneFrame fallback={<ColonFallback active={active} tumor={tumor} />} camera={[0, 0, 9]} fov={34}>
      <group position={[0, 0.65, 0]}>
        <ColonModel active={active} onSelect={interactive ? onSelect : undefined} labels={interactive} tumor={tumor} />
      </group>
      <OrbitControls
        enablePan={false} enableZoom={false} enableDamping
        minPolarAngle={1.1} maxPolarAngle={2.05} minAzimuthAngle={-0.9} maxAzimuthAngle={0.9}
      />
    </SceneFrame>
  );
}
