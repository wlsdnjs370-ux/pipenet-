"""Render the supplied SDF before/after display framing; no hydraulic edits."""
from __future__ import annotations

import json
from pathlib import Path
import sys
from types import SimpleNamespace
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'core'), str(ROOT/'pipenet_converter/src')]

from routes.module_f.merge import bake_combined_iso
from src.pipenet_converter.render.framing import reference_display_scale
from pipenet_converter.sdf_compat import prepare_sdf


def main() -> None:
    """Write a clearly marked historical display-only preview, never deploy it."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.patches import Circle, Polygon

    out = ROOT/'data/module_f_display_review_20260920'
    out.mkdir(exist_ok=True)
    original = Path('C:/Users/admin/Downloads/module_f_merged_iso (5).sdf')
    snapshot = json.loads((ROOT/'data/merge_boundary_review_20260920/live_before.json').read_text(encoding='utf-8'))
    rows = [dict(label=n['label'], x=n['xyz'][0]*1000, y=n['xyz'][1]*1000,
                 elevation=n['xyz'][2]) for n in snapshot['merge_editor']['nodes']]
    tree = ET.parse(original)
    by = {n.get('label'): n for n in tree.findall('.//Nodes/Node')}
    assert all(n['label'] in by and abs(n['elevation']-float(by[n['label']].get('elevation'))) < 1e-6
               for n in rows), 'Historical snapshot does not match this SDF'
    plan = [n['label'] for n in rows if n['label'].isdigit() and int(n['label']) >= 10]
    machine = [n['label'] for n in rows if n['label'].startswith('m')] + ['1']
    model = SimpleNamespace(nodes=rows, machine_room_plan_edges=[])
    projected = bake_combined_iso(dict(combined=model, system_layout='physical_xy'))[0]
    xs, ys = [n['x'] for n in projected], [n['y'] for n in projected]
    cx, cy = (min(xs)+max(xs))/2, (min(ys)+max(ys))/2
    old_scale = 3000/max(max(xs)-min(xs), max(ys)-min(ys))
    new_scale = reference_display_scale((n['x'], n['y']) for n in rows if n['label'] in plan)
    max_error = 0
    for n in projected:
        pos = by[n['label']].find('Position')
        for key, center in [('x', cx), ('y', cy)]:
            max_error = max(max_error, abs(float(pos.get(key))-(n[key]-center)*old_scale))
            pos.set(key, f'{(n[key]-center)*new_scale:.9g}')
    assert max_error < .02, f'Historical display differs: {max_error}'
    revised = out/'historical_display_only_NOT_recalculated.sdf'
    tree.write(revised, encoding='utf-8', xml_declaration=True)
    prepare_sdf(revised, nozzle_reference_labels=plan)

    def signature(path):
        r = ET.parse(path).getroot()
        for parent in r.iter():
            for child in parent.findall('Position'):
                parent.remove(child)
        return [(e.tag, sorted(e.attrib.items()), (e.text or '').strip())
                for tag in ('Nodes','Links') for e in r.find('.//'+tag).iter()]
    assert signature(original) == signature(revised), 'Calculation values changed'
    plt.rcParams['font.family'] = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False
    fig, axes = plt.subplots(2, 2, figsize=(16, 12), dpi=130)
    for col, path in enumerate((original, revised)):
        root = ET.parse(path).getroot()
        xy = {n.get('label'): (float(n.find('Position').get('x')), float(n.find('Position').get('y')))
              for n in root.findall('.//Nodes/Node')}
        for row, labels in enumerate((plan, machine)):
            ax = axes[row, col]
            labels = set(labels)
            points = [xy[k] for k in labels]
            for pipe in root.findall('.//Pipe'):
                a,b = pipe.get('input'), pipe.get('output')
                if a in labels and b in labels:
                    ax.plot([xy[a][0],xy[b][0]], [xy[a][1],xy[b][1]], c='#354757', lw=.65)
            for k in labels:
                ax.add_patch(Circle(xy[k], radius=9, color='#192a35'))
            for nozzle in root.findall('.//Nozzle'):
                a,b = nozzle.get('input'), nozzle.get('output')
                if a in labels:
                    x,y = xy[b]
                    up = 1 if y > xy[a][1] else -1
                    ax.plot([xy[a][0],x], [xy[a][1],y], c='#354757', lw=.65)
                    ax.add_patch(Polygon([(x-15,y-up*16), (x+15,y-up*16), (x,y+up*16)],
                                         fill=False, edgecolor='#192a35', lw=.7))
                    points.append((x,y))
            xmin,xmax = min(p[0] for p in points),max(p[0] for p in points)
            ymin,ymax = min(p[1] for p in points),max(p[1] for p in points)
            margin=max(xmax-xmin,ymax-ymin)*.04
            ax.set_xlim(xmin-margin,xmax+margin);ax.set_ylim(ymin-margin,ymax+margin)
            ax.set_aspect('equal'); ax.axis('off')
            ax.set_title(('기존 통합 출력' if col==0 else '평면도 기준 배율 유지')+' · '+('평면도' if row==0 else '기계실'), fontsize=14)
    fig.suptitle('통합 SDF 표시 개선 — 계산값·연결은 그대로',fontsize=20)
    fig.text(.5,.018,'첨부 SDF의 좌표 기반 비교 · PipeNet 실제 화면 캡처는 아님 · 이전 계산값으로 만든 표시 검토용',ha='center',fontsize=11,color='#526777')
    fig.tight_layout(rect=(0,.04,1,.96))
    fig.savefig(out/'display_comparison.png', facecolor='white')
    report = dict(original_scale=old_scale, revised_scale=new_scale,
                  enlargement=new_scale/old_scale, historical_projection_error=max_error,
                  calculation_unchanged=True, preview=str(out/'display_comparison.png'))
    (out/'report.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(json.dumps(report))


if __name__ == '__main__':
    main()
