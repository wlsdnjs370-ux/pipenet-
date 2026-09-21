"""Adapt explicit assumed floor heights to extracted schematic section paths."""
from __future__ import annotations

from copy import deepcopy
import math

from src.pipenet_converter.graph.elevation import uniform_floor_section


def apply_assumed_floor_height(riser: dict, floor_labels: list,
                              height_m: float) -> dict:
    """Recalculate Z/rise/length, preserving CAD coordinates and loss elements.

    A single unambiguous floor-label column supplies the section calibration.
    Drawing X remains the existing horizontal measure; Y is a section height,
    never an additional plan bearing. No project-specific layers are assumed.
    """
    columns: dict[int, list[tuple[int, float]]] = {}
    for x, y, floor, _name in floor_labels:
        if floor != 99:  # Roof labels need a real level, not the legacy sentinel.
            columns.setdefault(round(x), []).append((floor, float(y)))
    if not columns:
        raise ValueError("층 표기가 없어 층고를 환산할 수 없습니다. 층 기준선을 확인하세요.")
    count = max(len({f for f, _ in marks}) for marks in columns.values())
    candidates = [marks for marks in columns.values()
                  if len({f for f, _ in marks}) == count]
    profiles = [uniform_floor_section(marks, float(height_m)) for marks in candidates]
    if any(profile != profiles[0] for profile in profiles[1:]):
        raise ValueError("서로 다른 층 기준열이 있어 자동 선택할 수 없습니다.")
    profile = profiles[0]
    result = deepcopy(riser)
    at = {str(n['label']): n for n in result['nodes']}
    anchor = at[str(result['av_node_label'])]
    datum = profile.at(float(anchor['y']))
    for node in at.values():
        node['inferred_elevation'] = node.get('elevation', 0)
        node['elevation'] = round(profile.at(float(node['y'])) - datum, 3)
        node['elev_source'] = 'user_assumed_floor_height'
    for pipe in result['pipes']:
        a, b = at[str(pipe['in'])], at[str(pipe['out'])]
        pipe['inferred_length'] = pipe['length']
        pipe['inferred_elev'] = pipe.get('elev', 0)
        rise = round(b['elevation'] - a['elevation'], 3)
        length = math.hypot((float(b['x']) - float(a['x'])) / 1000, rise)
        if length <= 0:
            raise ValueError(f"{pipe['label']}: 층고 환산 후 길이가 0입니다.")
        pipe['length'] = length
        pipe['elev'] = rise
        pipe['elev_source'] = 'user_assumed_floor_height'
        pipe['length_source'] = 'section_xz_assumed_height'
    result['system_coordinate_mode'] = 'section_xz'
    result['assumed_floor_height_m'] = float(height_m)
    result['elevation_notice'] = (
        f"층고 {height_m:g}m 가정 — 층 사이·최상층 밖 연결점은 도면 위치 비례 환산; "
        "실제 설치높이 미확인")
    result['total_pipe_length_m'] = sum(p['length'] for p in result['pipes'])
    result.setdefault('floor_matching', {}).update(
        floor_height_mm=height_m * 1000, height_source='user_assumed_floor_height')
    result['elev_sources'] = {
        'nodes': {'user_assumed_floor_height': len(at)},
        'pipes': {'user_assumed_floor_height': len(result['pipes'])}}
    return result
