import * as THREE from 'three';
import type { SectionKey } from '../content/slides';

export const SECTION_KEYS: SectionKey[] = ['ascending', 'transverse', 'descending', 'sigmoid'];
export const SECTION_LABEL: Record<SectionKey, string> = {
  ascending: 'Ascending', transverse: 'Transverse', descending: 'Descending', sigmoid: 'Sigmoid',
};
const v = (x: number, y: number, z = 0) => new THREE.Vector3(x, y, z);

/** Colon in anatomical position, seen from the front (patient's right = viewer's left). */
export const COLON_CURVES: Record<SectionKey, THREE.CatmullRomCurve3> = {
  ascending: new THREE.CatmullRomCurve3([v(-1.6, -1.2), v(-1.65, -0.5), v(-1.62, 0.3), v(-1.6, 1.0)]),
  transverse: new THREE.CatmullRomCurve3([v(-1.6, 1.0), v(-1.1, 0.95, 0.05), v(-0.5, 0.6), v(0, 0.5, 0.08), v(0.5, 0.62), v(1.1, 0.95, 0.05), v(1.6, 1.1)]),
  descending: new THREE.CatmullRomCurve3([v(1.6, 1.1), v(1.68, 0.4), v(1.64, -0.4), v(1.6, -0.95)]),
  sigmoid: new THREE.CatmullRomCurve3([v(1.6, -0.95), v(1.3, -1.25, 0.05), v(0.85, -1.05), v(0.5, -1.4, 0.05), v(0.2, -1.55), v(0.05, -1.95)]),
};
export const RECTUM = new THREE.CatmullRomCurve3([v(0.05, -1.95), v(0.02, -2.4, -0.04)]);

const tp = COLON_CURVES.descending.getPointAt(0.85);
export const TUMOR_POS: [number, number, number] = [tp.x + 0.2, tp.y, tp.z + 0.1];

export const LABEL_POS: Record<SectionKey, [number, number, number]> = {
  ascending: [-1.62, -0.1, 0.3], transverse: [0, 0.5, 0.35], descending: [1.66, 0, 0.3], sigmoid: [0.85, -1.05, 0.3],
};

/** Tube with gentle haustral pouching. */
export function makeTube(curve: THREE.CatmullRomCurve3, radius: number, haustra = true) {
  const tubular = Math.max(24, Math.round(curve.getLength() * 40));
  const radial = 20;
  const g = new THREE.TubeGeometry(curve, tubular, radius, radial, false);
  if (haustra) {
    const pos = g.attributes.position as THREE.BufferAttribute;
    const c = new THREE.Vector3();
    const p = new THREE.Vector3();
    const folds = Math.max(4, Math.round(curve.getLength() * 2.6));
    for (let i = 0; i < pos.count; i++) {
      const t = Math.floor(i / (radial + 1)) / tubular;
      curve.getPointAt(t, c);
      p.fromBufferAttribute(pos, i);
      const edge = Math.min(1, Math.min(t, 1 - t) * 8);
      const k = 1 + 0.1 * edge * Math.sin(t * folds * Math.PI * 2);
      p.sub(c).multiplyScalar(k).add(c);
      pos.setXYZ(i, p.x, p.y, p.z);
    }
    g.computeVertexNormals();
  }
  return g;
}
