import { OrbitControls } from '@react-three/drei';
import { useFrame } from '@react-three/fiber';
import * as THREE from 'three';
import { useMemo, useRef } from 'react';
import SceneFrame from './SceneFrame';

const NORMAL = new THREE.Color('#f2b6c4');
const ABN = new THREE.Color('#c9567b');
const BURG = new THREE.Color('#7a1f3d');
const COLS = 9, ROWS = 5, GAP = 0.42;
const rand = (i: number) => Math.abs(Math.sin(i * 12.9898 + 4.1) * 43758.5453) % 1;

interface Goal { y: number; s: number; c: THREE.Color }
function goal(ix: number, iz: number, seed: number, stage: number): Goal {
  const d = Math.hypot(ix, iz * 1.3);
  const g = Math.exp((-d * d) / 2.2);
  const c = NORMAL.clone();
  let y = 0, s = 1;
  if (stage === 1) { s = 0.7 + seed * 0.8; c.lerp(ABN, Math.min(1, g * 1.5 * seed + 0.1 * seed)); }
  if (stage === 2) { y = g * 0.85; s = 0.95 + seed * 0.35 * g; c.lerp(ABN, Math.min(1, g * 1.1)); }
  if (stage === 3) {
    y = g * 1.1 + (seed - 0.5) * 0.3 * g; s = 1 + g * seed * 0.7; c.lerp(BURG, Math.min(1, g * 1.3));
    if (d < 1.5 && seed > 0.72) { y = -0.42; s = 0.85; }
  }
  return { y, s, c };
}

function Cells({ stage }: { stage: number }) {
  const cells = useMemo(() => {
    const a: { ix: number; iz: number; seed: number }[] = [];
    let i = 0;
    for (let r = 0; r < ROWS; r++) for (let c = 0; c < COLS; c++, i++) a.push({ ix: c - (COLS - 1) / 2, iz: r - (ROWS - 1) / 2, seed: rand(i) });
    return a;
  }, []);
  const goals = useMemo(() => cells.map((c) => goal(c.ix, c.iz, c.seed, stage)), [cells, stage]);
  const refs = useRef<(THREE.Mesh | null)[]>([]);
  useFrame((_, dt) => {
    const k = 1 - Math.exp(-dt * 3.5);
    cells.forEach((_c, i) => {
      const m = refs.current[i];
      if (!m) return;
      const g = goals[i];
      m.position.y += (g.y - m.position.y) * k;
      m.scale.setScalar(m.scale.x + (g.s - m.scale.x) * k);
      (m.material as THREE.MeshStandardMaterial).color.lerp(g.c, k);
    });
  });
  return (
    <group>
      {cells.map((c, i) => (
        <mesh key={i} ref={(el) => { refs.current[i] = el; }} position={[c.ix * GAP, 0, c.iz * GAP]}>
          <sphereGeometry args={[0.17, 20, 14]} />
          <meshStandardMaterial color="#f2b6c4" roughness={0.5} />
        </mesh>
      ))}
      <mesh position={[0, -0.55, 0]}>
        <boxGeometry args={[4.3, 0.5, 2.4]} />
        <meshStandardMaterial color="#e9c3cc" transparent opacity={0.7} roughness={0.8} />
      </mesh>
    </group>
  );
}

function Fallback({ stage }: { stage: number }) {
  return (
    <svg viewBox="0 0 300 160" className="fallback-svg" role="img" aria-label="Cell progression">
      <rect x="20" y="110" width="260" height="30" rx="8" fill="#e9c3cc" />
      <ellipse cx="150" cy={105 - stage * 8} rx={30 + stage * 12} ry={12 + stage * 10} fill={stage === 3 ? '#7a1f3d' : stage === 2 ? '#c9567b' : '#f2b6c4'} />
    </svg>
  );
}

export default function PathoScene({ stage }: { stage: number }) {
  return (
    <SceneFrame fallback={<Fallback stage={stage} />} camera={[0, 3.1, 5.3]} fov={34}>
      <group position={[0, -0.2, 0]}><Cells stage={stage} /></group>
      <OrbitControls enablePan={false} enableZoom={false} enableDamping minPolarAngle={0.7} maxPolarAngle={1.5} minAzimuthAngle={-0.6} maxAzimuthAngle={0.6} />
    </SceneFrame>
  );
}
