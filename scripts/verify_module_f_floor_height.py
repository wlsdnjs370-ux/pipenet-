"""Reproduce the selected Daemyeong riser with the user's explicit 4 m assumption."""
from __future__ import annotations

import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'core')]

from routes.module_f.subdrawing import extract_system, parse_subdrawing


def main() -> None:
    """Save a review table; do not alter saved handwork or a live session."""
    sys.stdout.reconfigure(encoding='utf-8')
    source = ROOT / 'routes/제출용[최종]/1. 입력도면 대명동 단위세대 계통도.dxf'
    entities, _ = parse_subdrawing(source)
    riser = extract_system(entities, (-645097.8040654995, 180030.6808502593),
                           (-660681.0832206832, 136464.4665509026),
                           layer_filter={'HSP'}, assumed_floor_height_m=4)
    folder = ROOT / 'data/merge_boundary_review_20260920'
    (folder / 'system_4m_assumed.json').write_text(
        json.dumps(riser, ensure_ascii=False, indent=2), encoding='utf-8')
    summary = {
        'basis': riser['elevation_notice'],
        'source_elevation_m': riser['nodes'][0]['elevation'],
        'av_elevation_m': riser['nodes'][-1]['elevation'],
        'total_length_m': riser['total_pipe_length_m'],
        'pipes': [{k: p[k] for k in ('label', 'in', 'out', 'inferred_length',
                                    'inferred_elev', 'length', 'elev')}
                  for p in riser['pipes']]}
    (folder / 'floor_height_comparison.json').write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == '__main__':
    main()
