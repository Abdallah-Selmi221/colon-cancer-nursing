import { useThree } from '@react-three/fiber';
import gsap from 'gsap';
import * as THREE from 'three';
import { ReactNode, useEffect, useMemo, useRef } from 'react';
import ColonFallback from '../illustrations/ColonFallback';
import ColonModel from './ColonModel';
import SceneFrame from './SceneFrame';

const ghost = { color: '#efcdd6', transparent: true, opacity: 0.55, roughness: 0.7, depthWrite: false } as const;
const v = (x: number, y: number, z = 0) => new THREE.Vector3(x, y, z);

function SmallIntestine() {
  const geo = useMemo(() => {
    const pts: THREE.Vector3[] = [];
    [0.15, -0.12, -0.39, -0.66].forEach((y, r) => {
      for (let s = 0; s <= 8; s++) {
        const x = r % 2 === 0 ? -0.95 + s * 0.2375 : 0.95 - s * 0.2375;
        pts.push(v(x, y + 0.05 * Math.sin(s * 1.7), -0.12));
      }
    });
    return new THREE.TubeGeometry(new THREE.CatmullRomCurve3(pts, false, 'centripetal'), 320, 0.1, 10);
  }, []);
  return <mesh geometry={geo}><meshStandardMaterial {...ghost} /></mesh>;
}

function Esophagus() {
  const geo = useMemo(() => new THREE.TubeGeometry(new THREE.CatmullRomCurve3([v(0.05, 3.2, -0.1), v(0.2, 2.6, -0.1), v(0.45, 2.15, -0.1)]), 30, 0.09, 10), []);
  return <mesh geometry={geo}><meshStandardMaterial {...ghost} /></mesh>;
}

/** Slow camera push-in, once. Respects reduced motion. */
function Rig({ children }: { children: ReactNode }) {
  const group = useRef<THREE.Group>(null);
  const { camera } = useThree();
  useEffect(() => {
    if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) { camera.position.set(0, 0, 10.2); return; }
    camera.position.set(0, 0.8, 13.5);
    const a = gsap.to(camera.position, { z: 10.2, y: 0, duration: 3.2, ease: 'power2.out' });
    const b = group.current ? gsap.fromTo(group.current.rotation, { y: -0.6 }, { y: 0.12, duration: 3.6, ease: 'power2.out' }) : null;
    return () => { a.kill(); b?.kill(); };
  }, [camera]);
  return <group ref={group}>{children}</group>;
}

export default function DigestiveScene() {
  return (
    <SceneFrame fallback={<ColonFallback tumor={false} />} camera={[0, 0, 10.2]} fov={32}>
      <Rig>
        <group position={[0, -0.4, 0]}>
          <mesh position={[-0.95, 1.95, -0.15]} rotation={[0, 0, 0.2]} scale={[1, 0.5, 0.38]}>
            <sphereGeometry args={[1, 32, 24]} /><meshStandardMaterial {...ghost} />
          </mesh>
          <mesh position={[0.75, 1.85, -0.1]} rotation={[0, 0, -0.5]} scale={[0.6, 0.42, 0.34]}>
            <sphereGeometry args={[1, 32, 24]} /><meshStandardMaterial {...ghost} />
          </mesh>
          <Esophagus />
          <SmallIntestine />
          <ColonModel baseColor="#c8647f" />
        </group>
      </Rig>
    </SceneFrame>
  );
}
