"""Compare display-only geometry in the user's separate and merged SDFs."""
from __future__ import annotations

import json
from pathlib import Path
import xml.etree.ElementTree as ET

import numpy as np


def main() -> None:
    """Report scale/shear and symbol settings without modifying either file."""
    paths = [Path('C:/Users/admin/Downloads') / name for name in (
        '1. 입력도면 대명동 단위세대 평면도_수리계산입력 (2).sdf',
        'module_f_merged_iso (5).sdf')]
    roots = [ET.parse(p).getroot() for p in paths]
    positions = [{n.get('label'): tuple(float(n.find('Position').get(k)) for k in ('x','y'))
                  for n in r.findall('.//Nodes/Node')} for r in roots]
    pairs = [(label, str(int(label)+9)) for label in positions[0]
             if label.isdigit() and str(int(label)+9) in positions[1]]
    a = np.array([positions[0][x] for x,y in pairs])
    b = np.array([positions[1][y] for x,y in pairs])
    design = np.column_stack((a, np.ones(len(a))))
    fit, *_ = np.linalg.lstsq(design, b, rcond=None)
    err = np.linalg.norm(design @ fit-b, axis=1)
    out = dict(matched_nodes=len(pairs), affine=fit.tolist(),
               max_residual=float(max(err)), median_residual=float(np.median(err)),
               worst=[(pairs[i], float(err[i])) for i in np.argsort(err)[-10:]],
               singular_scales=np.linalg.svd(fit[:2], compute_uv=False).tolist(),
               plan_bbox=[np.ptp(a,axis=0).tolist(), np.ptp(b,axis=0).tolist()])
    uniform = np.ptp(b[:, 0]) / np.ptp(a[:, 0])
    a_scaled = (a - a.min(axis=0)) * uniform + b.min(axis=0)
    nearest = np.linalg.norm(a_scaled[:, None, :] - b[None, :, :], axis=2).min(axis=1)
    out['geometric_comparison_independent_of_node_renumbering'] = dict(
        uniform_scale=uniform, nearest_median=float(np.median(nearest)),
        nearest_max=float(max(nearest)), within_1_unit=int(sum(nearest < 1)))
    for tag in ('Display-options','Node-schemes'):
        out[tag] = [ET.tostring(r.find('.//'+tag), encoding='unicode') for r in roots]
    out['counts'] = [{tag: len(r.findall('.//'+tag)) for tag in ('Node','Pipe','Nozzle')}
                     for r in roots]
    print(json.dumps(out, ensure_ascii=True, indent=2))


if __name__ == '__main__':
    main()
