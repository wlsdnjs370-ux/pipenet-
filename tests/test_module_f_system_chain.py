"""[오너 2026-09-22 · 그림 42~44] 계통도 여러 장 — «＋ 계통도 추가».

연결 축: 평면도 — 계통도 1 — 계통도 2 — … — 계통도 n — 기계실.
계통도 k 의 ① 점(기계실 쪽 끝)과 계통도 k+1 의 ② 점(평면도 쪽 끝)은 같은 한 노드다.

확인하는 것:
  1. 칸 — 더하기 · 이름(계통도 1 · 2) · 지우기 · 칸 사이 상태 독립 · 세션 전역
  2. 잇기 — 이름(s2…) · 평행이동 · 앞 계통도 기준 표고 · 급수원 옮기기 · PIPENET 라벨 규칙
  3. 결합 — 지금 길 그대로(기계실은 계통도 n 에, 평면도는 계통도 1 에) · 산출까지
  4. 계통도가 한 장이면 종전과 같다(같은 객체)
  5. 빈 계통도 칸은 결합 전에 멈추고 말한다
"""
from copy import deepcopy
import math
from pathlib import Path
import re
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
for path in (ROOT, ROOT / 'core'):
    sys.path.insert(0, str(path))

from routes.module_f import jobs
from routes.module_f.merge import MergeError, bake_combined_iso, merge_network
from routes.module_f.slots import (
    SESSION_KEYS, SYSTEM_MAX, _check_slot_kind, _slot_active, _slot_add_system,
    _slot_capture, _slot_remove_system, _slot_state, _slot_switch, slot_label,
    slot_role, system_kinds)
from routes.module_f.system_chain import chain_label, chain_risers
from test_module_f_system_layout import heads, system
from test_module_f_shared_nodes import machine_room
from test_module_f_network_editor import session

PIPENET_LABEL = re.compile(r"[A-Za-z][A-Za-z0-9]*|[0-9]+")


def second():
    """계통도 2 — 다른 도면(좌표계가 다름). ① 펌프(1) → ② 계통도 1 과 만나는 점(10)."""
    pts = {'1': (100000, 50000, 0.0), 'n2': (102000, 50000, 0.0), '10': (102000, 53000, 1.5)}
    return dict(extracted_from='dxf', av_node_label='10', input_node_label='1',
                nodes=[dict(label=k, x=x, y=y, elevation=z,
                            io_node='Input' if k == '1' else 'No')
                       for k, (x, y, z) in pts.items()],
                pipes=[dict(label='r1', **{'in': '1', 'out': 'n2'}, length=2, elev=0, dia=100, c=120),
                       dict(label='r2', **{'in': 'n2', 'out': '10'}, length=1.5, elev=1.5,
                            dia=100, c=120)],
                fittings=[dict(pipe='r2', type='elbow', count=1)])


def _sess():
    sess = jobs._new_session(key='chain-fixture')
    return sess


# ─────────────────────────────────────────────── 1. 칸
def test_계통도_칸을_더하고_이름은_자리로_매긴다():
    s = _sess()
    try:
        assert _slot_state(s)['slots'][1]['label'] == '계통도'     # 한 장일 때는 종전 이름
        k2 = _slot_add_system(s)
        k3 = _slot_add_system(s)
        assert (k2, k3) == ('system2', 'system3')
        st = _slot_state(s)
        assert [x['kind'] for x in st['slots']] == ['plan', 'system', 'system2', 'system3', 'machineroom']
        assert [x['label'] for x in st['slots']] == ['평면도', '계통도 1', '계통도 2', '계통도 3', '기계실']
        assert [x['removable'] for x in st['slots']] == [False, False, True, True, False]
        assert {x['role'] for x in st['slots'][1:4]} == {'system'}
        assert st['can_add_system'] is True
        # 가운데 칸을 지우면 뒤 칸의 이름이 당겨진다(칸 id 는 그대로).
        _slot_remove_system(s, 'system2')
        assert slot_label(s, 'system3') == '계통도 2'
        assert system_kinds(s) == ['system', 'system3']
        # 지운 번호는 다시 쓴다 — 새 칸은 기계실 쪽 끝에 붙는다.
        assert _slot_add_system(s) == 'system2'
        assert system_kinds(s) == ['system', 'system3', 'system2']
        assert slot_label(s, 'system2') == '계통도 3'
    finally:
        jobs._SESSIONS.pop(s['id'], None)


def test_칸끼리_상태가_섞이지_않고_목록은_세션_전역이다():
    s = _sess()
    try:
        s['key'] = '평면.dxf'
        kind = _slot_add_system(s)
        _slot_switch(s, 'system')
        s['key'], s['riser'] = '계통1.dxf', {'nodes': [1]}
        _slot_switch(s, kind)
        assert s.get('key') is None and s.get('riser') is None, '새 칸은 비어 있어야 한다'
        s['key'], s['riser'] = '계통2.dxf', {'nodes': [2]}
        assert 'system_extra' in SESSION_KEYS and 'system_extra' not in _slot_capture(s)
        _slot_switch(s, 'system')
        assert (s['key'], s['riser']) == ('계통1.dxf', {'nodes': [1]})
        assert s['system_extra'] == [kind], '칸 목록이 칸을 따라가면 안 된다'
        _slot_switch(s, 'plan')
        assert s['key'] == '평면.dxf'
        assert slot_role(kind) == 'system' and _check_slot_kind(kind, s) == kind
        with pytest.raises(ValueError):
            _check_slot_kind(kind)            # 세션 없이는 더한 칸 이름이 안 통한다
    finally:
        jobs._SESSIONS.pop(s['id'], None)


def test_보고_있는_칸을_지우면_앞_계통도로_옮기고_계통도1은_못_지운다():
    s = _sess()
    try:
        kind = _slot_add_system(s)
        _slot_switch(s, kind)
        s['key'] = '계통2.dxf'
        assert _slot_remove_system(s, kind) == 'system'
        assert _slot_active(s) == 'system' and kind not in (s.get('slots') or {})
        assert s.get('key') is None
        for bad in ('system', 'plan', 'machineroom', 'system9', None):
            with pytest.raises(ValueError):
                _slot_remove_system(s, bad)
        for _ in range(SYSTEM_MAX - 1):
            _slot_add_system(s)
        assert _slot_state(s)['can_add_system'] is False
        with pytest.raises(ValueError, match=f'{SYSTEM_MAX}장'):
            _slot_add_system(s)
    finally:
        jobs._SESSIONS.pop(s['id'], None)


# ─────────────────────────────────────────────── 2. 잇기
def test_한_장이면_받은_것을_그대로_돌려준다():
    one = system(True)
    assert chain_risers([one]) is one


def test_계통도_2를_옮겨_공통_노드에_잇는다():
    r1, r2 = system(True), second()
    b1, b2 = deepcopy(r1), deepcopy(r2)
    got = chain_risers([r1, r2], ['계통도 1', '계통도 2'])
    assert (r1, r2) == (b1, b2), '원본 칸의 추출 결과를 건드리면 안 된다'
    at = {n['label']: n for n in got['nodes']}
    # 이름: 계통도 1 은 그대로, 계통도 2 는 s2…, 계통도 2 의 ② 점은 계통도 1 의 ① 점이 된다.
    assert set(at) == {'1', 'n2', 'n3', '10', 's21', 's2n2'}
    assert [p['label'] for p in got['pipes']] == ['r1', 'r2', 'r3', 's2r1', 's2r2']
    last = got['pipes'][-1]
    assert (last['in'], last['out']) == ('s2n2', '1')
    assert got['fittings'][-1] == dict(pipe='s2r2', type='elbow', count=1)
    # 급수원은 계통도 2 의 ① 점으로 옮겨 간다.
    assert got['input_node_label'] == 's21' and got['av_node_label'] == '10'
    assert at['s21']['io_node'] == 'Input' and at['1']['io_node'] == 'No'
    assert sum(n.get('io_node') == 'Input' for n in got['nodes']) == 1
    # 평행이동(돌리기·늘리기 없음) — 계통도 2 의 ② 점이 계통도 1 의 ① 점 자리에 온다.
    dx, dy = 0 - 102000, 0 - 53000
    assert (at['s2n2']['x'], at['s2n2']['y']) == (102000 + dx, 50000 + dy)
    assert (at['s21']['x'], at['s21']['y']) == (100000 + dx, 50000 + dy)
    # 표고 — 앞(평면도 쪽) 계통도 기준: 계통도 2 전체가 −1.5 m 옮겨 간다.
    assert at['s21']['elevation'] == pytest.approx(-1.5)
    assert at['1']['elevation'] == 0
    # 배관 길이·낙차는 그대로다.
    assert [p['length'] for p in got['pipes'][3:]] == [2, 1.5]
    assert got['chain']['joints'] == ['1']
    assert got['chain']['shifts'] == [{'name': '계통도 2', 'dz_m': -1.5}]
    assert '계통도 2 표고 -1.500 m 맞춤' in got['chain']['step']
    labels = [n['label'] for n in got['nodes']] + [p['label'] for p in got['pipes']]
    assert all(PIPENET_LABEL.fullmatch(str(x)) for x in labels), labels


def test_세_장은_앞에서부터_차례로_잇는다():
    r3 = second()
    got = chain_risers([system(True), second(), r3])
    assert got['input_node_label'] == 's31'
    assert got['chain']['joints'] == ['1', 's21']
    at = {n['label']: n for n in got['nodes']}
    assert {'s31', 's3n2'} <= set(at) and 's310' not in at
    assert at['s31']['elevation'] == pytest.approx(-3.0)
    assert chain_label(3, 'r2') == 's3r2'


def test_빈_칸이나_끝점_없는_계통도는_잇지_않고_말한다():
    with pytest.raises(MergeError, match='계통도 2 의 경로가 아직 없습니다'):
        chain_risers([system(True), None])
    broken = second()
    broken['input_node_label'] = 'zz'
    with pytest.raises(MergeError, match='① 기계실 쪽 끝'):
        chain_risers([system(True), broken])
    template = second()
    template['extracted_from'] = 'template'
    with pytest.raises(MergeError, match='도면에서'):
        chain_risers([system(True), template])


# ─────────────────────────────────────────────── 3. 결합 · 산출
def test_이은_계통도로_지금_결합이_그대로_선다(tmp_path):
    chained = chain_risers([system(True), second()], ['계통도 1', '계통도 2'])
    got = merge_network(heads(), riser=chained, machineroom=machine_room(),
                        mode='hsp_pump')
    c = got['combined']
    assert got['attached'] and got['pump_junction'] == 's21', '기계실은 계통도 n 의 ① 점에 붙는다'
    assert got['chain_joints'] == ['1']
    assert any('계통도 2장 잇기' in s for s in got['steps'])
    assert got['checks']['components'] == 1
    assert {'s21', 's2n2', '1', 'n2'} <= set(got['parts']['system'])
    inputs = [n['label'] for n in c.nodes if n.get('io_node') == 'Input']
    assert inputs == ['m1']
    # 실제 길이 배치 — 모든 배관이 선언 길이만큼 떨어져 있다.
    at = {n['label']: n for n in c.nodes}
    for p in c.pipes:
        a, b = at[p['in']], at[p['out']]
        assert math.dist((a['x'] / 1000, a['y'] / 1000, a['elevation']),
                         (b['x'] / 1000, b['y'] / 1000, b['elevation'])) == pytest.approx(p['length'], abs=1e-6)
    from routes.module_f.emit import emit_merged
    files = emit_merged(c, tmp_path, iso_nodes=bake_combined_iso(got)[0],
                        display_reference_labels=got['parts']['plan'])
    assert files.get('kfp') and files.get('sdf_iso'), files['warnings']
    text = Path(files['sdf']).read_text(encoding='utf-8')
    assert 'label="s2r1"' in text and 'label="s2r2"' in text


def test_한_장_결합은_종전과_같다():
    one = merge_network(heads(), riser=system(True), mode='lsp_gravity')
    again = merge_network(heads(), riser=chain_risers([system(True)]), mode='lsp_gravity')
    assert one['combined'].__dict__ == again['combined'].__dict__
    assert one['chain_joints'] == [] and one['steps'] == again['steps']


# ─────────────────────────────────────────────── 4. 화면 API
def _client():
    from test_module_f_merge_view import _client as client
    return client()


def test_API_로_칸을_더하고_바꾸고_지운다():
    c = _client()
    s = jobs._new_session()
    try:
        r = c.post('/api/module-f/slot/add-system', json={'sid': s['id']})
        d = r.get_json()
        assert r.status_code == 200 and d['added'] == 'system2' and d['added_label'] == '계통도 2'
        assert [x['kind'] for x in d['slots']] == ['plan', 'system', 'system2', 'machineroom']
        r = c.post('/api/module-f/slot/switch', json={'sid': s['id'], 'kind': 'system2'})
        assert r.status_code == 200 and r.get_json()['active'] == 'system2'
        d = c.get(f"/api/module-f/sub/state?sid={s['id']}").get_json()
        assert d['kind'] == 'system2' and d['extracted'] is False
        d = c.get(f"/api/module-f/merge/state?sid={s['id']}").get_json()
        assert d['order'] == ['plan', 'system', 'system2', 'machineroom']
        assert d['system_missing'] == ['계통도 1', '계통도 2']
        assert d['labels']['system2'] == '계통도 2 입상관'
        r = c.post('/api/module-f/slot/remove-system', json={'sid': s['id'], 'kind': 'system'})
        assert r.status_code == 400
        r = c.post('/api/module-f/slot/remove-system', json={'sid': s['id'], 'kind': 'system2'})
        d = r.get_json()
        assert r.status_code == 200 and d['active'] == 'system' and d['removed_label'] == '계통도 2'
        assert [x['kind'] for x in d['slots']] == ['plan', 'system', 'machineroom']
        d = c.get(f"/api/module-f/merge/state?sid={s['id']}").get_json()
        assert d['order'] == ['plan', 'system', 'machineroom'] and d['system_missing'] == []
    finally:
        jobs._SESSIONS.pop(s['id'], None)


def test_빈_계통도_칸이_있으면_결합을_멈추고_말한다():
    from test_module_f_network_editor import design
    c = _client()
    s = jobs._new_session()
    try:
        s['design'] = design()
        s['supply_mode'] = 'lsp_gravity'
        s['slots']['system']['riser'] = system(True)
        _slot_add_system(s)
        r = c.post('/api/module-f/merge/build', json={'sid': s['id']})
        assert r.status_code == 400
        assert '계통도 2 의 경로가 아직 없습니다' in r.get_json()['message']
    finally:
        jobs._SESSIONS.pop(s['id'], None)


def test_결합_잡은_계통도들을_이어_결합하고_화면은_공통_노드를_이음매로_그린다(monkeypatch, tmp_path):
    from routes.module_f import network_edit as ne
    from routes.module_f.api_merge import rebuild_merged
    from test_module_f_network_editor import design
    monkeypatch.setattr(ne, 'HISTORY_DIR', tmp_path / 'history')
    c = _client()
    s = jobs._new_session(key='chain-merge')
    try:
        s['design'] = design()
        s['supply_mode'] = 'hsp_pump'
        s['slots']['system']['riser'] = system(True)
        kind = _slot_add_system(s)
        s['slots'][kind]['riser'] = second()
        summary = rebuild_merged(s, persist_overrides=False)
        got = s['merged']
        assert got['chain_joints'] == ['1'] and got['pump_junction'] is None
        assert any('계통도 2장 잇기' in x for x in summary['steps'])
        inputs = [n['label'] for n in got['combined'].nodes if n.get('io_node') == 'Input']
        assert inputs == ['s21']
        view = c.get(f"/api/module-f/merge/preview?sid={s['id']}").get_json()['view']
        anchors = {n['label'] for n in view['nodes'] if n.get('anchor')}
        assert {'10', '1'} <= anchors
        # 칸의 추출 결과는 그대로다(잇기는 사본에서).
        assert s['slots'][kind]['riser'] == second()
    finally:
        jobs._SESSIONS.pop(s['id'], None)


def test_단면_계통도와_섞이면_단면_쪽_y_를_펴서_잇는다():
    base = system(True)
    base['system_coordinate_mode'] = 'section_xz'
    mixed = chain_risers([base, second()])
    assert 'system_coordinate_mode' not in mixed
    at = {n['label']: n for n in mixed['nodes']}
    assert {at[k]['y'] for k in ('1', 'n2', 'n3', '10')} == {0.0}
    # 계통도 2 는 평면 좌표 그대로 옮겨 온다(② 점이 계통도 1 의 ① 점 자리).
    assert (at['s2n2']['x'], at['s2n2']['y']) == (102000 - 102000, 50000 - 53000)
    both = second()
    both['system_coordinate_mode'] = 'section_xz'
    kept = chain_risers([base, both])
    assert kept['system_coordinate_mode'] == 'section_xz'
    assert {n['label']: n for n in kept['nodes']}['n3']['y'] == 4000


# ─────────────────────────────────────────────── 5. 통합 화면 — 밑그림 · 공통 노드
#   [오너 2026-09-22 · 그림 45·46] 계통도 칸 수만큼 밑그림 줄이 나고, 공통 노드는
#   번호가 아니라 «어느 두 도면이 만나는 자리인지» 로 부른다.
def _world(edges):
    points = [point for edge in edges for point in (edge[:2], edge[2:])]
    xs, ys = zip(*points)
    return dict(bounds=dict(minx=min(xs), miny=min(ys), maxx=max(xs), maxy=max(ys)),
                bundles=[dict(segs=[v for edge in edges for v in edge], circles=[], arcs=[])])


def _two_system_session(sess):
    """평면도 · 계통도 1 · 계통도 2 · 기계실 — 밑그림(원도면 도형)까지 갖춘 세션."""
    from routes.module_f.api_merge import rebuild_merged
    from test_module_f_merge_underlays import drawings
    drawings(sess)                                   # 평면도 · 계통도 · 기계실 밑그림
    sess['slots']['system']['riser'] = system(True)  # 두 점으로 뽑은 경로라야 잇는다
    kind = _slot_add_system(sess)
    sess['slots'][kind].update(
        key='system2-source', riser=second(),
        world=_world([[100000, 50000, 102000, 50000], [102000, 50000, 102000, 53000]]))
    sess['supply_mode'] = 'hsp_pump'
    rebuild_merged(sess, persist_overrides=False)
    return kind


def test_공통_노드는_어느_두_도면이_만나는지로_부른다():
    from routes.module_f.system_chain import joint_names
    chained = chain_risers([system(True), second()], ['계통도 1', '계통도 2'])
    got = merge_network(heads(), riser=chained, machineroom=machine_room(), mode='hsp_pump')
    assert joint_names(got) == {'10': '평면도-계통도1', '1': '계통도1-계통도2',
                                's21': '계통도2-기계실'}
    one = merge_network(heads(), riser=system(True), machineroom=machine_room(),
                        mode='hsp_pump')
    assert joint_names(one) == {'10': '평면도-계통도', '1': '계통도-기계실'}  # 한 장이면 칸 이름 그대로
    alone = merge_network(heads(), riser=system(True), mode='lsp_gravity')
    assert joint_names(alone) == {'10': '평면도-계통도'}          # 기계실이 없으면 그 자리도 없다


@pytest.mark.parametrize('iso', [False, True])
def test_계통도마다_밑그림_줄이_나고_앞뒤_공통_노드에_맞춘다(session, iso):
    from routes.module_f.merge import bake_combined_plan
    from routes.module_f.underlays import reference_layers
    from test_module_f_merge_underlays import transform
    kind = _two_system_session(session)
    got = session['merged']
    before = deepcopy(vars(got['combined']))
    rows = reference_layers(session, got, iso=iso, geometry=True)
    assert [r['kind'] for r in rows] == ['plan', 'system', kind, 'machineroom']
    assert [r['label'] for r in rows] == ['평면도', '계통도 1', '계통도 2', '기계실']
    assert all(r['available'] and r['groups'] for r in rows), [r['reason'] for r in rows]
    assert rows[2]['bounds'] == session['slots'][kind]['world']['bounds']
    assert got['system_layout'] == 'physical_xy'
    nodes = bake_combined_iso(got)[0] if iso else bake_combined_plan(got)[0]
    at = {str(n['label']): (n['x'], n['y']) for n in nodes}
    for row, (av, inp, av_to, inp_to) in zip(
            rows[1:3], [((5000, 4000), (0, 0), '10', '1'),
                        ((102000, 53000), (100000, 50000), '1', 's21')]):
        a, b, c, d, _, _ = row['matrix']
        assert b == c == 0 and a == d and a != 0          # 기울이지도 찌그러뜨리지도 않는다
        # ② 점은 앞 공통 노드에 그대로 앉고, ① 점은 뒤 공통 노드 «높이» 에 맞는다.
        assert transform(row['matrix'], av) == pytest.approx(at[av_to], abs=1e-6)
        assert transform(row['matrix'], inp)[1] == pytest.approx(at[inp_to][1], abs=1e-6)
    assert '계통도1-계통도2' in rows[1]['note'] and '계통도2-기계실' in rows[2]['note']
    assert vars(got['combined']) == before                # 밑그림은 배관망을 건드리지 않는다


def test_결합_뒤에_더한_계통도_칸은_다시_결합하라고_말한다(session):
    from routes.module_f.underlays import reference_layers
    _two_system_session(session)
    third = _slot_add_system(session)                     # 결합은 그대로 두고 칸만 더한다
    session['slots'][third].update(key='system3-source', riser=second(),
                                   world=_world([[0, 0, 1000, 0]]))
    rows = {r['kind']: r for r in reference_layers(session, session['merged'], iso=True)}
    assert rows[third]['available'] is False
    assert '다시 결합' in rows[third]['reason']
    assert rows['system']['available'] and rows['machineroom']['available']


def test_화면이_공통_노드_이름을_받아_그린다(monkeypatch, tmp_path):
    from routes.module_f import network_edit as ne
    from routes.module_f.api_merge import rebuild_merged
    from test_module_f_network_editor import design
    monkeypatch.setattr(ne, 'HISTORY_DIR', tmp_path / 'history')
    c = _client()
    s = jobs._new_session(key='chain-joint-names')
    try:
        s['design'] = design()
        s['supply_mode'] = 'hsp_pump'
        s['slots']['system']['riser'] = system(True)
        kind = _slot_add_system(s)
        s['slots'][kind]['riser'] = second()
        rebuild_merged(s, persist_overrides=False)
        d = c.get(f"/api/module-f/merge/preview?sid={s['id']}").get_json()
        at = {n['label']: n for n in d['view']['nodes']}
        assert at['10']['joint'] == '평면도-계통도1'
        assert at['1']['joint'] == '계통도1-계통도2'
        assert d['counts']['joints'] == ['평면도-계통도1', '계통도1-계통도2']
        assert all('joint' not in n for n in d['view']['nodes'] if not n.get('anchor'))
        # 검사줄도 같은 이름으로 말한다(«기준점 10» 대신).
        st = c.get(f"/api/module-f/merge/state?sid={s['id']}").get_json()
        assert st['checks']['anchor_joint'] == '평면도-계통도1'
    finally:
        jobs._SESSIONS.pop(s['id'], None)
