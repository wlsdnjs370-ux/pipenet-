"""Preserve geometry and hydraulic rows while removing whole-drawing work from confirmation."""
from __future__ import annotations

import copy
import itertools
import random
import sys
from pathlib import Path

import pytest
from flask import Flask

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from routes.module_f.common import _boot
_boot()
from services.cad_import.convert.planar import _find_x_crossings


def exhaustive_crossings(pos, edges):
    """Independent all-pairs oracle retaining the pre-change predicates."""
    seg = [(a, b, pos[a], pos[b]) for a, b in edges if a in pos and b in pos]
    out = []
    def side(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    for (a, b, p, q), (c, d, r, s) in itertools.combinations(seg, 2):
        if {a, b} & {c, d}:
            continue
        if not ((side(r,s,p)>0) != (side(r,s,q)>0)
                and (side(p,q,r)>0) != (side(p,q,s)>0)):
            continue
        den = (q[0]-p[0])*(s[1]-r[1]) - (q[1]-p[1])*(s[0]-r[0])
        if abs(den) < 1e-9:
            continue
        t = ((r[0]-p[0])*(s[1]-r[1]) - (r[1]-p[1])*(s[0]-r[0])) / den
        out.append({'a': (a,b), 'b': (c,d), 'at': (p[0]+t*(q[0]-p[0]),p[1]+t*(q[1]-p[1]))})
    return sorted(out, key=lambda row: (row['a'], row['b']))


@pytest.mark.parametrize('offset', [0.0, -200.0, 2000.0])
def test_crossing_index_matches_all_pairs(offset):
    rng = random.Random(293)
    pos = {i: (offset+rng.uniform(-30,30), rng.uniform(-30,30)) for i in range(100)}
    # Grid boundaries, shared nodes, coincident endpoints, and huge diagonals.
    pos.update({100:(-2,0),101:(2,0),102:(0,-2),103:(0,2),
                104:(0,0),105:(0,0),106:(-100000,-100000),107:(100000,100000)})
    edges = [(i,i+1) for i in range(0,108,2)] + [(100,103),(104,101)]
    before = copy.deepcopy((pos,edges))
    assert _find_x_crossings(pos,edges) == exhaustive_crossings(pos,edges)
    assert (pos,edges) == before


def test_separated_metre_segments_do_not_trigger_all_pairs():
    pos = {2*i:(3.0*i,0.0) for i in range(400)}
    pos.update({2*i+1:(3.0*i+1.0,1.0) for i in range(400)})
    edges = [(2*i,2*i+1) for i in range(400)]
    calls = []
    previous = sys.getprofile()
    def count(frame, event, arg):
        if event == 'call' and frame.f_code.co_name == '_side':
            calls.append(1)
    try:
        sys.setprofile(count)
        assert _find_x_crossings(pos,edges) == []
    finally:
        sys.setprofile(previous)
    assert len(calls) < 400  # Old 2000-m grid performs 319,200 side calls.


def test_physical_degrees_equal_exact_unique_edge_count():
    from services.cad_import.design.restrict import corridor_topology
    rng=random.Random(829)
    pts=[(rng.uniform(-125,125),rng.uniform(-125,125)) for _ in range(100)]
    edges=[(rng.randrange(100),rng.randrange(100)) for _ in range(250)]
    edges += [(0,0),(0,1),(0,1),(1,1000)]  # Internal, duplicate and missing-coordinate nodes.
    origin=(-115.0,125.0)
    # Nearness/display snapping is not evidence of one physical junction.
    unique={tuple(sorted((a,b))) for a,b in edges if a!=b}
    used={n for edge in edges for n in edge}
    expected={str(vid):sum(vid in edge for edge in unique) for vid in used}
    refs={str(vid):vid for vid in used}
    got=corridor_topology({'pts':pts,'edges':edges},
                         {'origin_mm':origin,'node_ref':refs},
                         {'nodes_meta_runtime':dict.fromkeys(refs,{}),'pipe_data':{}})
    assert got['phys']==expected


@pytest.fixture
def design_client(tmp_path, monkeypatch):
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    from services.cad_import.pipeline import handoff
    from routes.module_f import api_design, jobs
    from services.cad_import.design.worst import worst_k_heads
    old = handoff.import_write_root()
    handoff.set_write_root(str(tmp_path))
    board = EditBoard('__speed_test__',
                      [(0,0),(1000,0),(3000,0),(5000,0),(3000,1000),(5000,1000),(9000,9000)],
                      {(0,1),(1,2),(2,3),(2,4),(3,5)},
                      [(3000,1000,50),(5000,1000,50),(9000,9000,50)])
    for disk in board.disks:
        board.set_head_kind(disk, '상향식')
    board.sources = [0]
    board.valves = [1]
    es = EditSession(board, key='__speed_test__', out_dir=str(tmp_path))
    sess = jobs._new_session()
    sess.update(edit=es, key=es.key, worst_k=2,
                worst=worst_k_heads(board.pts,board.edges,board.hnodes,board.sources,
                                    k=2,head_xy=board.disks))
    def run(session, phase, fn):
        session['job'] = {'state':'run', 'phase':phase}
        result = fn()
        session['job'].update(state='done', result=result)
    monkeypatch.setattr(api_design, '_run_job', run)
    monkeypatch.setattr(api_design, '_dia_texts', lambda sess: [])
    app = Flask(__name__)
    app.config['TESTING'] = True
    api_design.register(app, UPLOAD_DIR=tmp_path)
    try:
        yield app.test_client(), sess
    finally:
        jobs._SESSIONS.pop(sess['id'], None)
        handoff.set_write_root(old)


def confirm(client, sess):
    response = client.post('/api/module-f/design/build', json={'sid':sess['id'], 'k':2})
    assert response.status_code == 200, response.get_json()
    return sess['job']['result']


def test_confirmation_never_requests_whole_network(design_client, monkeypatch):
    from routes.module_f import attach
    client, sess = design_client
    monkeypatch.setattr(attach, 'wet_heads', lambda *a, **k: pytest.fail('whole-network probe during confirmation'))
    before = copy.deepcopy(sess['worst'])
    result = confirm(client,sess)
    assert result['ok'], result
    assert result['summary']['counts']['nozzles'] == 2
    assert result['summary']['diagnostics_state'] == 'pending'
    assert result['summary']['excluded_heads'] is None
    assert result['summary']['total_heads'] == 3
    assert 'unattached' not in sess['design']['marks']
    assert sess['worst'] == before


@pytest.mark.parametrize('wet,shared', [(set(),set()), ({0},set()), ({0,1},{1})])
def test_missing_selected_head_still_blocks_without_backfill(design_client, monkeypatch, wet, shared):
    from routes.module_f import attach
    client, sess = design_client
    monkeypatch.setattr(attach, 'picked_heads_wet', lambda *a, **k: {
        'ok':True,'wet':wet,'shared':shared,'reason':{},'total':2,'dropped':2-len(wet)})
    before = copy.deepcopy(sess['worst'])
    result = confirm(client,sess)
    assert not result['ok']
    assert '배관에 붙지 않습니다' in result['error']
    assert sess['worst'] == before
    assert not sess.get('design')


def hydraulic_rows(sess):
    tables = sess['design']['tables']
    return copy.deepcopy([getattr(tables, name) for name in ('nodes','pipes','nozzles','fittings','equipment')])


def test_diagnosis_preserves_rows_and_matches_previous_whole_probe(design_client, monkeypatch, tmp_path):
    from routes.module_f import attach
    from services.cad_import.design.emit import emit_design_sdf
    client,sess = design_client
    assert confirm(client,sess)['ok']
    rows = hydraulic_rows(sess)
    out = tmp_path/'comparison.sdf'
    emit_design_sdf(sess['design']['tables'], out)
    original_sdf = out.read_bytes()
    worst = copy.deepcopy(sess['worst'])
    response = client.post('/api/module-f/design/diagnose', json={'sid':sess['id']})
    assert response.status_code == 200, response.get_json()
    result = sess['job']['result']
    assert result['ok'], result
    assert result['summary']['diagnostics_state'] == 'done'
    assert 'unattached' in sess['design']['marks']
    assert hydraulic_rows(sess) == rows
    emit_design_sdf(sess['design']['tables'], out)
    assert out.read_bytes() == original_sdf
    assert sess['worst'] == worst
    monkeypatch.setattr(attach, 'design_probe', lambda s,e,picked,**kw: attach.wet_heads(s,e,**kw))
    assert confirm(client,sess)['ok']
    assert hydraulic_rows(sess) == rows
    emit_design_sdf(sess['design']['tables'], out)
    assert out.read_bytes() == original_sdf


def test_diagnosis_rejects_stale_tables(design_client):
    client,sess = design_client
    assert confirm(client,sess)['ok']
    sess['worst']['heads'] = [0]
    response = client.post('/api/module-f/design/diagnose', json={'sid':sess['id']})
    assert response.status_code == 409


def test_diagnostic_cache_detects_middle_coordinate_change(design_client, monkeypatch):
    from routes.module_f import attach
    from services.cad_import.design import restrict
    _, sess = design_client
    calls = []
    def probe(*args, **kwargs):
        calls.append(1)
        return {'ok': True, 'wet': {0,1}, 'dropped':1}
    monkeypatch.setattr(restrict, 'attachable_heads', probe)
    es=sess['edit']
    attach.wet_heads(sess,es)
    attach.wet_heads(sess,es)
    assert len(calls)==1
    es.board.pts[1]=(1001,0)  # Same count and same last three points.
    attach.wet_heads(sess,es)
    assert len(calls)==2
