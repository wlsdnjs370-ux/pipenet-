"""Inspect the supplied Daemyeong extraction inputs without changing them."""
from __future__ import annotations

from pathlib import Path
import sys
import json
import math

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'core')]

from routes.module_f.subdrawing import parse_subdrawing, path_graph
from remote30_prototype import _auto_pipe_layer_filter, _extract_floor_labels, _collapse_collinear_nodes


def main() -> None:
    """Report raw long runs near the exported suspect system pipes."""
    source = ROOT / 'routes' / '제출용[최종]' / '1. 입력도면 대명동 단위세대 계통도.dxf'
    entities, _ = parse_subdrawing(source)
    layers = sorted(_auto_pipe_layer_filter(entities))
    print('layers', layers)
    sys.stdout.reconfigure(encoding='utf-8')
    for layer in layers:
        graph, lengths, stats = path_graph(entities, layer_filter={layer})
        print('layer', layer, 'nodes', len(graph))
        if layer == 'HSP':
            from remote30_prototype import extract_system_path
            expected = json.loads((ROOT/'data/merge_boundary_review_20260920/live_before.json').read_text(encoding='utf-8'))
            target = [p for p in expected['merge_editor']['pipes'] if p['label'].startswith('r')]
            # Endpoint windows follow the unique exported long horizontal run;
            # only inspect plausible vertices, not arbitrary drawing geometry.
            starts = [p for p in graph if p[0]>-650000 and 180000<p[1]<185000]
            ends = [p for p in graph if abs(p[0]+660681)<2 and 135000<p[1]<139000]
            for start in starts:
                for end in ends:
                    raw = extract_system_path(entities,start,end,layer_filter={layer})
                    if len(raw['pipes']) == len(target) and all(abs(p['length']-t['length'])<.002 for p,t in zip(raw['pipes'],target)):
                        out=ROOT/'data/merge_boundary_review_20260920/system_before.json'
                        out.write_text(json.dumps(raw,ensure_ascii=False,indent=2),encoding='utf-8')
                        print('EXACT MATCH',start,end)
                        for p in raw['pipes']:
                            if p['label'] in ('r1','r2','r3','r4','r5','r23'):
                                at={n['label']:n for n in raw['nodes']}
                                print(p,'NODES',at[p['in']],at[p['out']])
        seen = set()
        for a in graph:
            if len(graph[a]) == 2:
                continue
            for b in graph[a]:
                path = [a, b]
                while len(graph[path[-1]]) == 2:
                    nxt = next(n for n in graph[path[-1]] if n != path[-2])
                    if nxt in path:
                        break
                    path.append(nxt)
                pair = tuple(sorted((path[0], path[-1])))
                if pair in seen:
                    continue
                seen.add(pair)
                simple = _collapse_collinear_nodes(path, dict(lengths), angle_tol_deg=2)
                for u,v in zip(simple,simple[1:]):
                    if abs(abs(v[0]-u[0])-15583) < 50 or abs(v[1]-u[1])>30000:
                        print('run', u, v, 'raw vertices', len(path), 'simple', simple)


if __name__ == '__main__':
    main()
