import { ThreeEvent, useFrame } from '@react-three/fiber';
import { Html } from '@react-three/drei';
import * as THREE from 'three';
import { useMemo, useRef } from 'react';
import type { SectionKey } from '../content/slides';
import { COLON_CURVES, LABEL_POS, RECTUM, SECTION_KEYS, SECTION_LABEL, TUMOR_POS, makeTube } from './geometry';

const DIM = new THREE.Color('#ecc7d0');
const HI = new THREE.Color('#8a2145');

interface PartProps {
  geo: THREE.BufferGeometry; k: SectionKey; active: SectionKey | null; base: THREE.Color;
  onSelect?: (k: SectionKey) => void; position?: [number, number, number];
}

function Part({ geo, k, active, base, onSelect, position }: PartProps) {
  const ref = useRef<THREE.Mesh>(null);
  const target = active === null ? base : active === k ? HI : DIM;
  useFrame((_, dt) => {
    const m = ref.current?.material as THREE.MeshStandardMaterial | undefined;
    if (m) m.color.lerp(target, 1 - Math.exp(-dt * 6));
  });
  const click = onSelect ? (e: ThreeEvent<MouseEvent>) => { e.stopPropagation(); onSelect(k); } : undefined;
  return (
    <mesh
      ref={ref} geometry={geo} position={position} onClick={click}
      onPointerOver={onSelect ? (e) => { e.stopPropagation(); document.body.style.cursor = 'pointer'; } : undefined}
      onPointerOut={onSelect ? () => { document.body.style.cursor = 'auto'; } : undefined}
    >
      <meshStandardMaterial color={base} roughness={0.55} />
    </mesh>
  );
}

interface Props {
  active?: SectionKey | null; onSelect?: (k: SectionKey) => void; tumor?: boolean; baseColor?: string; labels?: boolean;
}

export default function ColonModel({ active = null, onSelect, tumor = false, baseColor = '#d98ba0', labels = false }: Props) {
  const geos = useMemo(
    () => Object.fromEntries(SECTION_KEYS.map((k) => [k, makeTube(COLON_CURVES[k], 0.26)])) as unknown as Record<SectionKey, THREE.BufferGeometry>,
    [],
  );
  const rectum = useMemo(() => makeTube(RECTUM, 0.2, false), []);
  const ball = useMemo(() => new THREE.SphereGeometry(0.262, 24, 16), []);
  const base = useMemo(() => new THREE.Color(baseColor), [baseColor]);
  const joints: [SectionKey, [number, number, number]][] = [
    ['ascending', [-1.6, -1.2, 0]], ['ascending', [-1.6, 1, 0]], ['descending', [1.6, 1.1, 0]], ['sigmoid', [1.6, -0.95, 0]],
  ];
  return (
    <group>
      {SECTION_KEYS.map((k) => <Part key={k} geo={geos[k]} k={k} active={active} base={base} onSelect={onSelect} />)}
      {joints.map(([k, p], i) => <Part key={i} geo={ball} k={k} active={active} base={base} onSelect={onSelect} position={p} />)}
      <mesh geometry={rectum}><meshStandardMaterial color="#e5b3bf" roughness={0.6} /></mesh>
      {tumor && (
        <group position={TUMOR_POS}>
          <mesh><icosahedronGeometry args={[0.2, 2]} /><meshStandardMaterial color="#7a1f3d" roughness={0.8} /></mesh>
          {([[0.14, 0.12, 0.1], [0.16, -0.1, 0.08], [-0.02, 0.2, 0.12], [0.05, -0.2, 0.1], [0.2, 0, 0.16]] as [number, number, number][]).map((p, i) => (
            <mesh key={i} position={p}><icosahedronGeometry args={[0.1, 1]} /><meshStandardMaterial color="#8f2748" roughness={0.8} /></mesh>
          ))}
        </group>
      )}
      {labels && SECTION_KEYS.map((k) => (
        <Html key={k} position={LABEL_POS[k]} center zIndexRange={[10, 0]}>
          <button className={`tag ${active === k ? 'on' : ''}`} onClick={() => onSelect?.(k)}>{SECTION_LABEL[k]}</button>
        </Html>
      ))}
    </group>
  );
}
