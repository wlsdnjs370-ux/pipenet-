"""Compare saved source-cache area extraction without modifying a live session."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> None:
    """Write before/after source identities and a plan-view review PNG."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cache', type=Path, required=True)
    parser.add_argument('--source', type=int, required=True)
    parser.add_argument('--zone', type=float, nargs=4, required=True)
    parser.add_argument('--mode', choices=('loop', 'grid'), default='loop')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.design.flow import network_for_board
    from services.cad_import.design.preserved import expand_preserved
    from src.pipenet_converter.graph.area_selection import select_area_heads
    from src.pipenet_converter.graph.area_network import select_area_network
    data = json.loads(args.cache.read_text(encoding='utf-8'))['data']
    board = _board_from_data('area-audit', data)
    board.sources = [args.source]
    board.valves = [args.source]
    board.network_mode = args.mode
    network = network_for_board(board)
    selected = select_area_heads(board.disks, [args.zone], network.reference)
    if selected.disconnected:
        raise ValueError(f'Disconnected area heads: {selected.disconnected}')
    before = network.selected_edges(selected.heads)
    scope = select_area_network(network, board.pts, selected.heads, selected.zones)
    got = expand_preserved(board, selected.heads, source_edges=tuple(scope.edges), area_scope=scope.report(network))
    args.output.mkdir(parents=True, exist_ok=True)
    report = {
        'source_cache': str(args.cache), 'source_node': args.source, 'zones': selected.zones,
        'heads': list(selected.heads), 'source_network': network.report(),
        'before': {'edges': len(before), 'length_m': sum(network.reference.lengths_mm[e] for e in before)/1000},
        'after': scope.report(network),
        'model': {'nodes': len(got['kfp']['nodes_meta_runtime']), 'pipes': len(got['kfp']['pipe_data']),
                  'nozzles': sum(n['type_id']=='head' for n in got['kfp']['nodes_meta_runtime'].values()),
                  'cycle_rank': got['cycle_rank']},
        'selected_source_edges': [list(e) for e in sorted(scope.edges)],
    }
    (args.output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    (args.output/'selected-network.json').write_text(json.dumps(network.extraction(
        board.pts, selected_heads=selected.heads, source_edges=scope.edges), ensure_ascii=False), encoding='utf-8')
    (args.output/'selected-model.json').write_text(json.dumps(got['kfp'], ensure_ascii=False), encoding='utf-8')
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection
    from matplotlib.patches import Rectangle
    fig, axes = plt.subplots(1, 2, figsize=(12, 7), facecolor='#070b0e')
    for ax, edges, title in zip(axes, (before, scope.edges), ('Before: all exterior paths', 'After: region + supply')):
        ax.set_facecolor('#070b0e')
        ax.add_collection(LineCollection([[board.pts[a],board.pts[b]] for a,b in network.edges],
                                         colors='#23313b', linewidths=.45))
        ax.add_collection(LineCollection([[board.pts[a],board.pts[b]] for a,b in edges],
                                         colors='#eef6fc', linewidths=1.5))
        x0,y0,x1,y1 = args.zone
        ax.add_patch(Rectangle((x0,y0),x1-x0,y1-y0,fill=False,edgecolor='#38bdf8',linewidth=1.1))
        ax.scatter([board.disks[h][0] for h in selected.heads], [board.disks[h][1] for h in selected.heads],
                   s=22,edgecolors='#ff4b57',facecolors='none',zorder=4)
        ax.scatter([board.pts[args.source][0]], [board.pts[args.source][1]],c='#38bdf8',s=25,zorder=4)
        ax.autoscale(); ax.set_aspect('equal'); ax.axis('off')
        length = sum(network.reference.lengths_mm[e] for e in edges)/1000
        ax.set_title(f'{title}\n{len(edges)} edges / {length:.2f} m / {len(selected.heads)} heads',color='#e6edf3',fontsize=11)
    fig.tight_layout()
    fig.savefig(args.output/'comparison.png',dpi=150,facecolor=fig.get_facecolor())
    plt.close(fig)
    print(json.dumps({k:report[k] for k in ('heads','source_network','before','after','model')}, ensure_ascii=False))


if __name__ == '__main__':
    main()
