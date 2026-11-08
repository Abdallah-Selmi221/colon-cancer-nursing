import { Canvas } from '@react-three/fiber';
import { Component, ReactNode, Suspense } from 'react';

let cached: boolean | undefined;
function webglOK() {
  if (cached !== undefined) return cached;
  try {
    const c = document.createElement('canvas');
    cached = !!(c.getContext('webgl2') || c.getContext('webgl'));
  } catch { cached = false; }
  return cached;
}

class Boundary extends Component<{ fallback: ReactNode; children: ReactNode }, { failed: boolean }> {
  state = { failed: false };
  static getDerivedStateFromError() { return { failed: true }; }
  componentDidCatch() { /* fallback is shown instead */ }
  render() { return this.state.failed ? this.props.fallback : this.props.children; }
}

interface Props { children: ReactNode; fallback: ReactNode; camera?: [number, number, number]; fov?: number }

/** Canvas wrapper: lights, transparent background, and a flat fallback when WebGL is missing or fails. */
export default function SceneFrame({ children, fallback, camera = [0, 0, 8], fov = 32 }: Props) {
  if (!webglOK()) return <div className="scene">{fallback}</div>;
  return (
    <div className="scene">
      <Boundary fallback={fallback}>
        <Suspense fallback={null}>
          <Canvas dpr={[1, 1.75]} camera={{ position: camera, fov }} gl={{ antialias: true, alpha: true }}>
            <hemisphereLight args={['#fff4f6', '#d9a0b0', 1.1]} />
            <directionalLight position={[3, 4, 5]} intensity={1.7} />
            <directionalLight position={[-4, -1, 3]} intensity={0.5} color="#ffd6e0" />
            {children}
          </Canvas>
        </Suspense>
      </Boundary>
    </div>
  );
}
