import { Html, OrbitControls } from '@react-three/drei';
import * as THREE from 'three';
import SceneFrame from './SceneFrame';

function Fallback() {
  return (
    <svg viewBox="0 0 300 200" className="fallback-svg" role="img" aria-label="Colon cross-section">
      <rect x="10" y="40" width="280" height="120" rx="30" fill="#c47e92" />
      <rect x="10" y="52" width="280" height="96" rx="24" fill="#e9a8b8" />
      <circle cx="150" cy="130" r="30" fill="#7a1f3d" />
      <text x="40" y="100" fontSize="13" fill="#3b2430">Normal tissue</text>
      <text x="112" y="182" fontSize="13" fill="#3b2430">Malignant growth</text>
    </svg>
  );
}

const side = THREE.DoubleSide;

export default function CrossSectionScene() {
  return (
    <SceneFrame fallback={<Fallback />} camera={[0, 1.2, 5.6]} fov={38}>
      <group rotation={[0, 0, Math.PI / 2]}>
        <mesh>
          <cylinderGeometry args={[1, 1, 4.2, 64, 1, true, Math.PI / 2, Math.PI]} />
          <meshStandardMaterial color="#e9a8b8" side={side} roughness={0.6} />
        </mesh>
        <mesh>
          <cylinderGeometry args={[1.16, 1.16, 4.2, 64, 1, true, Math.PI / 2, Math.PI]} />
          <meshStandardMaterial color="#c47e92" side={side} roughness={0.7} />
        </mesh>
        {[1, -1].map((s) => (
          <mesh key={s} position={[s * 1.08, 0, 0]}>
            <boxGeometry args={[0.16, 4.2, 0.03]} /><meshStandardMaterial color="#b86a80" />
          </mesh>
        ))}
        <group position={[0, 0.1, -0.78]}>
          <mesh><icosahedronGeometry args={[0.3, 2]} /><meshStandardMaterial color="#7a1f3d" roughness={0.85} /></mesh>
          {([[0.28, 0.1, 0.05], [-0.26, 0.16, 0.02], [0.1, -0.3, 0.04], [-0.14, -0.26, 0.06], [0.05, 0.32, 0.06], [0.32, -0.2, 0.0]] as [number, number, number][]).map((p, i) => (
            <mesh key={i} position={p}><icosahedronGeometry args={[0.15, 1]} /><meshStandardMaterial color="#8f2748" roughness={0.85} /></mesh>
          ))}
        </group>
      </group>
      <Html position={[1.5, 0.8, 0.2]} center><span className="tag static">Normal tissue</span></Html>
      <Html position={[-0.1, -1.35, 0.2]} center><span className="tag static hot">Malignant growth</span></Html>
      <OrbitControls enablePan={false} enableZoom={false} enableDamping minPolarAngle={0.9} maxPolarAngle={1.9} minAzimuthAngle={-0.7} maxAzimuthAngle={0.7} />
    </SceneFrame>
  );
}
