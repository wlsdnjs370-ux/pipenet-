"""[오너 2026-09-21] 통합망 계통도 세로관 삭제 — 「제거 후 상단과 하단 배관을 연결할까요?」

- 예(delete_join): 배관을 지우고 위·아래 노드를 붙인다. 급수원(기계실) 쪽 표고는
  그대로, 반대쪽(평면도 쪽)이 지운 배관 높이만큼 옮겨진다 — 입상관이 한 칸 짧아진다.
- 아니오(delete_cut): 배관만 지운다. 통합망이 끊겼다는 문구를 하단에 띄우고 산출을 막는다.
- 노드 점은 실제 크기(지름 0.5 m)라 줌에 따라 커지고 작아진다.
"""
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest
from flask import Flask

from core.network_editor import EditError, Network, apply_edit
from routes.module_f import api_merge, network_edit as ne
from routes.module_f.merge import SPLIT_MESSAGE, merge_network, network_pieces
from test_module_f_network_editor import client, session, post, state, cmd  # noqa: F401 (fixtures)

ROOT = Path(__file__).resolve().parents[1]
for _path in (ROOT, ROOT/'core'):  # 결합 엔진(core/remote30_full_network)을 찾는 경로
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

# 급수원 1(11.4 m) → n2(7.6) → n3(3.8) → n4(가로 3 m) → 기준점 10(0.0 m): 세로관 r1·r2·r4, 가로관 r3.
Z = {'1': 11.4, 'n2': 7.6, 'n3': 3.8, 'n4': 3.8, '10': 0.0}
X = {'1': 0, 'n2': 0, 'n3': 0, 'n4': 3000, '10': 3000}


def riser():
    order = ['1', 'n2', 'n3', 'n4', '10']
    nodes = [dict(label=k, x=X[k], y=0, elevation=Z[k], io_node='Input' if k == '1' else 'No') for k in order]
    pipes = []
    for i, (a, b) in enumerate(zip(order, order[1:])):
        dz, dx = Z[b]-Z[a], abs(X[b]-X[a])/1000
        pipes.append(dict(label=f'r{i+1}', **{'in': a, 'out': b}, length=round((dx*dx+dz*dz)**.5, 6),
                          elev=dz, dia=100, c=120))
    return dict(extracted_from='dxf', av_node_label='10', input_node_label='1', nodes=nodes, pipes=pipes)


def plan():
    return SimpleNamespace(
        nodes=[dict(label='1', x=3000, y=0, elevation=0, io_node='Input'),
               dict(label='2', x=5000, y=0, elevation=0, io_node='No'),
               dict(label='3', x=5000, y=0, elevation=-.3, io_node='No')],
        pipes=[dict(label='P1', **{'in': '1', 'out': '2'}, length=2, elev=0, dia=25, c=120, type='KSD 3507'),
               dict(label='P2', **{'in': '2', 'out': '3'}, length=.3, elev=-.3, dia=25, c=120, type='KSD 3507')],
        nozzles=[dict(label='1', **{'in': '3', 'out': '@/1'}, status='1', lib='SP-HEAD', flow_m3s=80/60000)],
        fittings=[], equipment=[], meta=[])


def network():
    got = merge_network(plan(), riser=riser(), mode='lsp_gravity')
    assert got['system_layout'] == 'physical_xy'
    parts = got['parts']
    seam = set(parts['plan']) & (set(parts['system']) | set(parts['machineroom']))
    return got, Network.from_tables(got['combined'], protected=seam)


def z(net, label):
    return net.nodes[label].xyz[2]


def test_예_기계실쪽_표고는_그대로_평면도쪽이_한칸_옮겨진다():
    got, net = network()
    out, result = apply_edit(net, dict(op='delete_join', target='r2', keep=['10']), ne.catalog())
    # r2(n2→n3, 3.8 m) 삭제 → n3 이 n2 에 붙는다. 급수원 쪽(n2)의 라벨이 남는다.
    assert result == dict(kind='node', label='n2', counts=result['counts'])
    assert 'r2' not in out.pipes and 'n3' not in out.nodes
    assert (out.pipes['r3'].a, out.pipes['r3'].b) == ('n2', 'n4')
    assert len(out.nodes) == len(net.nodes)-1 and len(out.pipes) == len(net.pipes)-1
    # 급수원 쪽은 한 치도 안 움직인다.
    for k in ('1', 'n2'):
        assert out.nodes[k].xyz == net.nodes[k].xyz
    # 평면도 쪽(기준점 10·헤드 포함)은 전부 +3.8 m, XY 는 그대로.
    moved = [k for k in out.nodes if k not in ('1', 'n2')]
    assert moved and '10' in moved
    for k in moved:
        assert out.nodes[k].xyz[:2] == net.nodes[k].xyz[:2]
        assert z(out, k)-z(net, k) == pytest.approx(3.8)
    # 표에도 그대로 — 평면도 표고가 바뀌고 배관 높이차(elev)가 새 좌표를 따른다.
    rows = {r['label']: r for r in out.tables.nodes}
    assert rows['10']['elevation'] == pytest.approx(3.8)
    assert rows['1']['elevation'] == pytest.approx(11.4)
    assert {r['label']: r for r in out.tables.pipes}['r3']['elev'] == pytest.approx(0)
    assert network_pieces(out.tables) == 1
    assert len(net.nodes) == len(out.nodes)+1  # 미리보기는 원본을 안 건드린다


def test_기준점_바로_위_배관이면_기준점이_남는다():
    _, net = network()
    out, result = apply_edit(net, dict(op='delete_join', target='r4', keep=['10']), ne.catalog())
    assert result['label'] == '10' and 'n4' not in out.nodes
    assert (out.pipes['r3'].a, out.pipes['r3'].b) == ('n3', '10')
    assert z(out, '10') == pytest.approx(3.8) and z(out, '1') == pytest.approx(11.4)


def test_급수원_바로_아래_배관이면_급수원이_남는다():
    _, net = network()
    out, result = apply_edit(net, dict(op='delete_join', target='r1', keep=['10']), ne.catalog())
    assert result['label'] == '1' and 'n2' not in out.nodes
    assert out.nodes['1'].xyz == net.nodes['1'].xyz
    assert z(out, '10') == pytest.approx(3.8)


def test_높이차_없는_배관과_양끝을_다_남겨야_하는_배관은_붙이지_않는다():
    _, net = network()
    with pytest.raises(EditError, match='높이차가 없는'):
        apply_edit(net, dict(op='delete_join', target='r3', keep=['10']), ne.catalog())
    with pytest.raises(EditError, match='높이차가 없는'):
        apply_edit(net, dict(op='delete_cut', target='r3', keep=['10']), ne.catalog())
    short = riser()
    short['nodes'] = [short['nodes'][0], short['nodes'][-1]]
    short['nodes'][1]['x'] = 0
    short['pipes'] = [dict(label='r1', **{'in': '1', 'out': '10'}, length=11.4, elev=-11.4, dia=100, c=120)]
    got = merge_network(plan(), riser=short, mode='lsp_gravity')
    net = Network.from_tables(got['combined'], protected={'10'})
    with pytest.raises(EditError, match='모두 급수원'):
        apply_edit(net, dict(op='delete_join', target='r1', keep=['10']), ne.catalog())


def test_아니오는_배관만_지운다_좌표는_그대로_두_덩어리가_남는다():
    _, net = network()
    out, result = apply_edit(net, dict(op='delete_cut', target='r2', keep=['10']), ne.catalog())
    assert result['label'] == 'n2'
    assert 'r2' not in out.pipes and set(out.nodes) == set(net.nodes)
    assert all(out.nodes[k].xyz == net.nodes[k].xyz for k in net.nodes)
    assert network_pieces(out.tables) == 2
    # 끊긴 채로도 다시 이을 수 있다 — 두 끝은 같은 세로줄 위에 있다.
    joined, _ = apply_edit(out, dict(op='connect', target='n2', end='n3', schedule='KSD 3507', dn=100),
                           ne.catalog())
    assert network_pieces(joined.tables) == 1


# ─────────────────────────────────────── 통합 화면 API
def merged(sess):
    sess['slots'] = {'system': {'riser': riser()}}
    sess['supply_mode'] = 'lsp_gravity'
    api_merge.rebuild_merged(sess, persist_overrides=False)
    assert sess['merged']['system_layout'] == 'physical_xy'
    ne.ensure(sess, 'merge')


def ok(response):
    assert response.status_code == 200, response.json
    return response.json


def current(sess):
    return ne.ensure(sess, 'merge')['current']


def test_API_예_는_남길_노드를_기록하고_되돌리기로_복원된다(client, session):
    merged(session)
    before = {k: n.xyz for k, n in current(session).nodes.items()}
    r = ok(post(client, session, dict(op='delete_join', target='r2'), action='preview', scope='merge'))
    assert 'n3' not in {n['label'] for n in r['preview']['nodes']}
    assert 'n3' in current(session).nodes  # 미리보기는 저장하지 않는다
    r = ok(post(client, session, dict(op='delete_join', target='r2'), scope='merge'))
    assert r['selection']['label'] == 'n2'
    net = current(session)
    assert 'n3' not in net.nodes and z(net, '1') == pytest.approx(11.4) and z(net, '10') == pytest.approx(3.8)
    assert ne.ensure(session, 'merge')['commands'][-1]['keep'] == ['10']
    assert {str(n['label']) for n in session['merged']['combined'].nodes} == set(net.nodes)
    assert 'n3' not in session['merged']['parts']['system']
    ok(post(client, session, action='undo', scope='merge'))
    assert {k: n.xyz for k, n in current(session).nodes.items()} == before


def test_API_평면도_편집_뒤에도_세로관_붙이기가_유지된다(client, session):
    merged(session)
    ok(post(client, session, dict(op='delete_join', target='r2'), scope='merge'))
    ok(post(client, session, cmd(target='11'), scope='merge'))  # 평면도 쪽 편집 → 다시 결합
    net = current(session)
    assert 'n3' not in net.nodes and z(net, '10') == pytest.approx(3.8)


def test_API_계통도_세로관이_아니면_카드_동작을_거절한다(client, session):
    merged(session)
    r = post(client, session, dict(op='delete_join', target='P2'), action='preview', scope='merge')
    assert r.status_code == 409 and '계통도' in r.json['message']
    r = post(client, session, dict(op='delete_cut', target='r2'), action='preview')
    assert r.status_code == 409


def test_아니오_뒤에는_하단_경고와_산출_거절_되돌리면_풀린다(client, session, tmp_path):
    merged(session)
    assert api_merge._split_note(session) is None
    ok(post(client, session, dict(op='delete_cut', target='r2'), scope='merge'))
    assert 'r2' not in current(session).pipes
    assert api_merge._split_note(session) == SPLIT_MESSAGE
    app = Flask(__name__)
    app.testing = True
    api_merge.register(app, UPLOAD_DIR=str(tmp_path/'uploads'))
    r = app.test_client().post('/api/module-f/merge/emit', json={'sid': session['id']})
    assert r.status_code == 409 and '연결되지 않아' in r.json['message']
    ok(post(client, session, action='undo', scope='merge'))
    assert api_merge._split_note(session) is None


def test_화면_카드와_노드_점():
    html = (ROOT/'templates'/'module_f.html').read_text(encoding='utf-8')
    editor = (ROOT/'static'/'module_f_editor.js').read_text(encoding='utf-8')
    main = (ROOT/'static'/'module_f.js').read_text(encoding='utf-8')
    assert '제거 후 상단과 하단 배관을 연결할까요?' in html
    assert all(f'id="{i}"' in html for i in ('ne-join-yes', 'ne-join-no', 'mg-split'))
    assert "op:join?'delete_join':'delete_cut'" in editor
    # 노드 점은 실제 길이(mm) × 화면 배율 — 줌 고정이 아니다.
    assert 'const MERGE_NODE_R_MM = 250;' in main
    assert 'const nodeR = MERGE_NODE_R_MM * S.view.scale;' in main
    assert 'S.mergeView.split' in main
