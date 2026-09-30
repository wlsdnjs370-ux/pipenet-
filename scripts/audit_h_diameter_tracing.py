"""Read-only saved-graph diameter audit; writes only the requested new report."""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def audit(key: str, historical: bool = False) -> dict:
    """Analyze a retained source graph without rebuilding caches or live sessions."""
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import _disp_cache_load, _board_from_data, load_edits
    from services.cad_import.design.flow import flow_for_board
    from routes.module_f.api_design import _dia_texts
    from routes.module_f.bore_context import build_bore_context
    data = _disp_cache_load(key)
    used_snapshot = False
    if data is None and historical:
        from services.cad_import.pipeline.disp_cache import _disp_cache_path
        data = json.loads(Path(_disp_cache_path(key)).read_text(encoding='utf-8'))['data']
        used_snapshot = True
    if data is None:
        raise ValueError('No current saved graph. No user cache was rebuilt.')
    board = _board_from_data(key, data)
    load_edits(board)
    if not board.sources:
        raise ValueError('The saved graph requires a supply source.')
    before = deepcopy((board.pts, board.edges, board.disks))
    session = {'key': key, 'design_settings': {'diameter_policy': 'drawing_first_v1'}}
    started = time.perf_counter()
    context = build_bore_context(session, board, flow_for_board(board), _dia_texts(session))
    if before != (board.pts, board.edges, board.disks):
        raise AssertionError('Diameter interpretation changed the source geometry.')
    reasons = Counter(r.get('reason', 'unknown') for r in context.decisions.values())
    sources = Counter(r.get('definition_source', 'unresolved') for r in context.decisions.values())
    return dict(key=key, historical_snapshot=used_snapshot, seconds=time.perf_counter()-started,
                summary=context.summary, reasons=dict(reasons), sources=dict(sources),
                annotations=[asdict(a) for a in session['_h_diameter_annotations']],
                annotation_audit=[asdict(a) for a in context.annotation_audit],
                decisions=[dict(edge=e, xy=[board.pts[n][:2] for n in e], evidence=r)
                           for e, r in sorted(context.decisions.items())])


def compare(before: dict, after: dict) -> dict:
    """Compare physical XY edges, not unstable integer IDs or annotation counts."""
    def indexed(report: dict) -> dict:
        rows = {tuple(sorted(tuple(round(v,6) for v in p) for p in r['xy'])):r
                for r in report['decisions']}
        if len(rows) != len(report['decisions']):
            raise ValueError('Coincident edges require explicit identity review before comparison.')
        return rows
    old,new = indexed(before),indexed(after)
    if set(old) != set(new):
        raise ValueError('The saved physical graphs differ; coverage is not directly comparable.')
    def value(row: dict) -> int | None:
        return row['evidence'].get('text_mm') if not row['evidence'].get('block_export') else None
    transitions: Counter = Counter()
    changes = []
    for key in sorted(old):
        a,b = value(old[key]),value(new[key])
        transitions[f'{"known" if a else "unknown"} -> {"known" if b else "unknown"}'] += 1
        if a == b:
            continue
        changes.append(dict(xy=key,before_mm=a,after_mm=b,
                            before=old[key]['evidence'],after=new[key]['evidence']))
    return dict(geometry_unchanged=True,physical_edges=len(old),transitions=dict(transitions),
                changed_known_values=sum(r['before_mm'] is not None and r['after_mm'] is not None for r in changes),
                before_summary=before['summary'],after_summary=after['summary'],changes=changes,
                limitation='Coverage diagnostic only; changed values require drawing review, not a certified accuracy measurement.')


def main() -> None:
    """Write a new JSON report; never silently replace a comparison baseline."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--key')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--historical', action='store_true')
    parser.add_argument('--compare', nargs=2, type=Path, metavar=('BEFORE','AFTER'))
    args = parser.parse_args()
    if args.output.exists():
        parser.error('Report exists; choose a new output path.')
    if args.compare:
        if args.key:
            parser.error('Use either --key or --compare.')
        result = compare(*(json.loads(p.read_text(encoding='utf-8')) for p in args.compare))
    elif args.key:
        result = audit(args.key, args.historical)
    else:
        parser.error('Supply --key or --compare.')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open('x', encoding='utf-8') as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2, allow_nan=False)
    fields=('geometry_unchanged','physical_edges','transitions','changed_known_values') if args.compare else ('seconds','summary','reasons','sources')
    print(json.dumps({k:result[k] for k in fields},ensure_ascii=False))


if __name__ == '__main__':
    main()
