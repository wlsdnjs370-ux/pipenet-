"""User-assumed floor heights must not turn section Y into phantom XY pipes."""
from copy import deepcopy
from pathlib import Path
import math
import xml.etree.ElementTree as ET

import pytest

from src.pipenet_converter.graph.elevation import uniform_floor_section
from routes.module_f.floor_elevation import apply_assumed_floor_height
from routes.module_f.subdrawing import extract_system, parse_subdrawing
from routes.module_f.merge import merge_network, bake_combined_iso
from test_module_f_system_layout import heads


def test_uneven_drawing_spacing_is_still_four_metres_per_floor():
    profile = uniform_floor_section([(30, 1000), (31, 3100), (32, 5946)], 4)
    assert profile.at(3100) - profile.at(1000) == 4
    assert profile.at(5946) - profile.at(3100) == 4
    assert profile.at(2050) - profile.at(1000) == 2
    assert profile.at(-50) - profile.at(1000) == -2
    assert profile.at(7369) - profile.at(5946) == 2


@pytest.mark.parametrize('height', [0, -1, math.nan, math.inf])
def test_invalid_height_does_not_silently_fall_back(height):
    with pytest.raises(ValueError):
        uniform_floor_section([(1, 0), (2, 1000)], height)


def test_conflicting_levels_missing_labels_and_basement_transition():
    with pytest.raises(ValueError):
        uniform_floor_section([(1, 1000), (2, 500)], 4)
    with pytest.raises(ValueError):
        uniform_floor_section([(1, 0)], 4)
    with pytest.raises(ValueError):
        apply_assumed_floor_height({}, [], 4)
    profile = uniform_floor_section([(-1, 0), (1, 1000)], 4)
    assert profile.at(1000) - profile.at(0) == 4


@pytest.fixture(scope='module')
def real_path():
    root = Path(__file__).resolve().parents[1]
    source = root / 'routes/제출용[최종]/1. 입력도면 대명동 단위세대 계통도.dxf'
    entities, _ = parse_subdrawing(source)
    return extract_system(entities, (-645097.8040654995, 180030.6808502593),
                          (-660681.0832206832, 136464.4665509026),
                          layer_filter={'HSP'}, assumed_floor_height_m=4)


def test_daemyeong_suspect_runs_are_vertical_with_explicit_assumed_elevations(real_path):
    assert len(real_path['pipes']) == 23
    at = {n['label']: n for n in real_path['nodes']}
    pipes = {p['label']: p for p in real_path['pipes']}
    assert pipes['r5']['length'] == 4
    assert pipes['r5']['elev'] == -4
    assert 4.4 < pipes['r23']['length'] < 4.42  # AV below the floor-label line.
    assert pipes['r4']['elev'] < 0  # Former nearest-floor quantization gave zero.
    assert at['10']['elevation'] == 0
    assert all(n['elev_source'] == 'user_assumed_floor_height' for n in at.values())
    before = deepcopy(real_path)
    merged = merge_network(heads(), riser=real_path, mode='lsp_gravity')
    nodes = {n['label']: n for n in merged['combined'].nodes}
    for p in real_path['pipes']:
        a, b = nodes[p['in']], nodes[p['out']]
        assert math.dist((a['x']/1000, a['y']/1000, a['elevation']),
                         (b['x']/1000, b['y']/1000, b['elevation'])) == pytest.approx(p['length'])
        if p['label'] != 'r3':
            assert (a['x'], a['y']) == pytest.approx((b['x'], b['y']))
    assert merged['pipe_parts']['r23'] == 'system'
    assert any('4m 가정' in s for s in merged['steps'])
    assert sum(n['label'] == '10' for n in nodes.values()) == 1
    assert real_path == before


def test_recalculated_values_survive_sdf_iso_and_kfp_export(real_path, tmp_path):
    from routes.module_f.emit import emit_merged
    merged = merge_network(heads(), riser=real_path, mode='lsp_gravity')
    files = emit_merged(merged['combined'], tmp_path, iso_nodes=bake_combined_iso(merged)[0])
    assert files.get('kfp'), files['warnings']
    for kind in ('sdf', 'sdf_iso'):
        tree = ET.parse(files[kind])
        pipes = {p.get('label'): p for p in tree.findall('.//Pipe')}
        assert float(pipes['r5'].get('length')) == 4
        assert float(pipes['r5'].get('rise')) == -4
        assert float(pipes['r23'].get('length')) == pytest.approx(-float(pipes['r23'].get('rise')))


def test_same_level_choice_ignores_assumed_floor_height():
    entities = [dict(t='L', l='PIPE', p=[0, 0, 0, 2100])]
    got = extract_system(entities, (0, 0), (0, 2100), layer_filter={'PIPE'},
                         elevation_mode='same_level', assumed_floor_height_m=4)
    assert got['pipes'][0]['length'] == 2.1
    assert got['pipes'][0]['elev'] == 0
    assert 'assumed_floor_height_m' not in got


def test_api_applies_explicit_height_and_invalid_input_preserves_result():
    from flask import Flask
    from routes.module_f import api_sub, jobs
    app = Flask(__name__)
    app.testing = True
    api_sub.register(app)
    sess = jobs._new_session(key='floor-height-fixture')
    sess.update(active='system', entities=[
        dict(t='L', l='PIPE', p=[0, 0, 0, 4200]),
        dict(t='T', l='TEXT', p=[-1000, 0], v='지상30층 SL'),
        dict(t='T', l='TEXT', p=[-1000, 2100], v='지상31층 SL'),
        dict(t='T', l='TEXT', p=[-1000, 4200], v='지상32층 SL')])
    body = dict(sid=sess['id'], pump_x=0, pump_y=4200, av_x=0, av_y=0,
                layers=['PIPE'], elevation_mode='drawing', assumed_floor_height_m=4)
    try:
        client = app.test_client()
        response = client.post('/api/module-f/system/extract', json=body)
        assert response.status_code == 200, response.json
        assert response.json['summary']['total_m'] == 8
        assert response.json['summary']['assumed_floor_height_m'] == 4
        assert '가정' in response.json['summary']['elevation_notice']
        before = deepcopy(sess['riser'])
        for invalid in ('bad', 0, -4):
            response = client.post('/api/module-f/system/extract',
                                   json={**body, 'assumed_floor_height_m': invalid})
            assert response.status_code == 400
            assert sess['riser'] == before
    finally:
        jobs._SESSIONS.pop(sess['id'], None)
