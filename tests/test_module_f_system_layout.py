"""Selected system routes retain bearings; reference drawings never steer pipes."""
from copy import deepcopy
import json
import math
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
for path in (ROOT, ROOT / 'core'):
    sys.path.insert(0, str(path))

from routes.module_f.merge import bake_combined_iso, bake_combined_plan, merge_network
from routes.module_f.system_layout import set_elevation_mode


def heads():
    return SimpleNamespace(
        nodes=[dict(label='1', x=1000, y=1000, elevation=0, io_node='Input'),
               dict(label='2', x=2000, y=1000, elevation=0, io_node='No'),
               dict(label='3', x=2000, y=1000, elevation=-.3, io_node='No')],
        pipes=[dict(label='P1', **{'in':'1','out':'2'}, length=1, elev=0, dia=25, c=120),
               dict(label='P2', **{'in':'2','out':'3'}, length=.3, elev=-.3, dia=25, c=120)],
        nozzles=[dict(label='1', **{'in':'3','out':'@/1'}, status='1', lib='SP-HEAD', flow_m3s=80/60000)],
        fittings=[], equipment=[], meta=[])


def system(vertical=False):
    points = [(0,0,0),(3000,0,0),(3000,4000,2 if vertical else 0),(5000,4000,2 if vertical else 0)]
    labels = ['1','n2','n3','10']
    lengths = [3, 2 if vertical else 4, 2]
    return dict(extracted_from='dxf', av_node_label='10', input_node_label='1',
                nodes=[dict(label=k,x=p[0],y=p[1],elevation=p[2],io_node='Input' if k=='1' else 'No')
                       for k,p in zip(labels,points)],
                pipes=[dict(label=f'r{i+1}', **{'in':a,'out':b},length=lengths[i],
                            elev=points[i+1][2]-points[i][2],dia=100,c=120)
                       for i,(a,b) in enumerate(zip(labels,labels[1:]))])


@pytest.mark.parametrize('vertical', [False,True])
@pytest.mark.parametrize('mode', ['hsp_pump','lsp_gravity'])
def test_selected_bearings_lengths_and_height_independent_of_supply_mode(vertical,mode):
    raw=system(vertical); before=deepcopy(raw)
    got=merge_network(heads(),riser=raw,mode=mode)
    assert got['system_layout']=='physical_xy'
    at={n['label']:n for n in got['combined'].nodes}
    assert at['n2']['x']-at['1']['x']==pytest.approx(3000)
    assert at['n2']['y']==at['1']['y']
    assert at['10']['x']-at['n3']['x']==pytest.approx(2000)
    assert at['n3']['y']-at['n2']['y']==pytest.approx(0 if vertical else 4000)
    iso={n['label']:n for n in bake_combined_iso(got)[0]}
    if vertical:
        assert iso['n3']['x']==iso['n2']['x']
        assert iso['n3']['y']-iso['n2']['y']==pytest.approx(2000)
    else:
        assert abs(iso['n2']['x']-iso['1']['x'])>2000
    assert raw==before
    for p in raw['pipes']:
        a,b=at[p['in']],at[p['out']]
        assert math.dist((a['x']/1000,a['y']/1000,a['elevation']),
                         (b['x']/1000,b['y']/1000,b['elevation']))==pytest.approx(p['length'])


def test_same_level_does_not_use_y_or_inferred_clamped_length_as_height():
    raw=system(True)
    raw['nodes'][0]['elevation']=16
    raw['pipes'][0]['length']=16
    measured={((0,0),(3000,0)):3000,((3000,0),(3000,4000)):4000,((3000,4000),(5000,4000)):2000}
    result=set_elevation_mode(raw,'same_level',measured)
    assert all(n['elevation']==0 for n in result['nodes'])
    assert [p['length'] for p in result['pipes']]==[3,4,2]
    assert all(p['elev']==0 for p in result['pipes'])
    assert raw['nodes'][0]['elevation']==16
    assert set_elevation_mode(raw,'drawing')['nodes']==raw['nodes']
    with pytest.raises(ValueError,match='실측 길이'):
        set_elevation_mode(raw,'same_level')


def test_same_floor_choice_reaches_real_extractor_and_keeps_selected_path():
    from routes.module_f.subdrawing import extract_system
    entities=[dict(t='L',l='PIPE',p=[0,0,3000,0]),dict(t='L',l='PIPE',p=[3000,0,3000,4000])]
    raw=extract_system(entities,(0,0),(3000,4000),layer_filter={'PIPE'},elevation_mode='same_level')
    assert raw['elevation_mode']=='same_level'
    assert sum(p['length'] for p in raw['pipes'])==pytest.approx(7)
    assert all(n['elevation']==0 for n in raw['nodes'])
    got=merge_network(heads(),riser=raw,mode='hsp_pump')
    at={n['label']:(n['x'],n['y']) for n in got['combined'].nodes}
    assert at['1'][0] != at['10'][0]


def test_export_reuses_physical_bearings_and_preserves_all_edges(tmp_path):
    from routes.module_f.emit import emit_merged
    got=merge_network(heads(),riser=system(True),mode='hsp_pump')
    output=emit_merged(got['combined'],tmp_path,iso_nodes=bake_combined_iso(got)[0])
    assert 'kfp' in output, output['warnings']
    data=json.loads(Path(output['kfp']).read_text(encoding='utf-8'))
    at={k:n['coords'] for k,n in data['nodes_meta_runtime'].items()}
    assert len(data['pipe_data'])==5
    first=next(p for p in data['pipe_data'].values() if p['start']=='N1')
    second=next(p for p in data['pipe_data'].values() if p['start']==first['end'])
    a,b=at[first['end']],at[second['end']]
    assert a[0]-at['N1'][0]==pytest.approx(3)
    assert a[:2]==b[:2]
    assert b[2]-a[2]==pytest.approx(2)
    assert Path(output['kfp']).read_bytes()==Path(output['kfp_iso']).read_bytes()


@pytest.mark.parametrize('iso',[False,True])
def test_upright_underlay_does_not_relayout_physical_network(iso):
    from routes.module_f.underlays import reference_layers
    raw=system(True)
    got=merge_network(heads(),riser=raw,mode='hsp_pump')
    sess=dict(active='plan',slots={'system':dict(key='mixed',riser=raw,world=dict(bundles=[dict(segs=[0,0,3000,0])]))})
    before=deepcopy(vars(got['combined']))
    refs=reference_layers(sess,got,iso=iso,geometry=True)
    ref=next(r for r in refs if r['kind']=='system')
    assert ref['available']
    a,b,c,d,_,_=ref['matrix']
    assert b==c==0 and a==d and a!=0
    assert vars(got['combined'])==before


@pytest.mark.parametrize('iso',[False,True])
def test_machine_reference_stays_attached_after_system_layout(session,iso):
    from test_module_f_merge_underlays import drawings, transform
    from routes.module_f.api_merge import rebuild_merged
    from routes.module_f.underlays import reference_layers
    drawings(session)
    session['slots']['system']['riser']=system(True)
    rebuild_merged(session,persist_overrides=False)
    got=session['merged']
    # 평면 보기는 계통도 세로관을 세운 자리에 그린다(오너 2026-09-21) — 밑그림은
    # 화면에 보이는 그 자리에 붙어야 한다.
    nodes=bake_combined_iso(got)[0] if iso else bake_combined_plan(got)[0]
    at={str(n['label']):(n['x'],n['y']) for n in nodes}
    refs={r['kind']:r for r in reference_layers(session,got,iso=iso)}
    assert transform(refs['machineroom']['matrix'],(7000,6000))==pytest.approx(at['1'])
    for label, point in [('m1',(4000,9000)),('m2',(7000,9000))]:
        assert transform(refs['machineroom']['matrix'],point)==pytest.approx(at[label],abs=1.1)


# Reuse the existing isolated edit-session fixture; no live session is touched.
from test_module_f_network_editor import session


# ── [오너 2026-09-21] 통합 평면 보기 — 계통도 세로관을 세로로 ─────────────────────
#
# physical_xy 는 계통도를 위에서 내려다본 실제 자리에 둔다. 위에서 보면 세로관은
# 길이가 0 이라 평면 보기에서 입상관이 한 점으로 접혔다(대명동 실측: 세로관 26개).
# 평면 보기만 (표고 − 기준점 표고)×1000 을 화면 세로로 더한다. 실제 표·KFP·아이소는
# 그대로다.

def test_plan_view_stands_vertical_pipes_up_and_keeps_level_pipes():
    got=merge_network(heads(),riser=system(True),mode='hsp_pump')
    before=deepcopy(vars(got['combined']))
    real={n['label']:n for n in got['combined'].nodes}
    plan={n['label']:n for n in bake_combined_plan(got)[0]}
    assert real['n2']['x']==real['n3']['x'] and real['n2']['y']==real['n3']['y']  # 위에서 보면 한 점
    # 세로관 n2→n3 (높이차 2 m) → 화면 세로 2000 mm
    assert plan['n3']['x']==pytest.approx(plan['n2']['x'])
    assert plan['n3']['y']-plan['n2']['y']==pytest.approx(2000)
    # 높이차 없는 배관은 도면 방향 그대로
    assert plan['n2']['x']-plan['1']['x']==pytest.approx(3000)
    assert plan['n2']['y']==pytest.approx(plan['1']['y'])
    assert plan['10']['x']-plan['n3']['x']==pytest.approx(2000)
    # 평면도와 기준점 10 은 한 치도 안 움직인다
    for label in got['parts']['plan']:
        assert (plan[label]['x'],plan[label]['y'])==(real[label]['x'],real[label]['y'])
    assert vars(got['combined'])==before


def test_plan_view_is_unchanged_for_same_level_and_template_layouts():
    measured={((0,0),(3000,0)):3000,((3000,0),(3000,4000)):4000,((3000,4000),(5000,4000)):2000}
    same=set_elevation_mode(system(True),'same_level',measured)
    sys.path.insert(0,str(ROOT/'tests'))
    from test_module_f_merge import _riser
    for got in (merge_network(heads(),riser=same,mode='hsp_pump'),
                merge_network(heads(),riser=_riser(),mode='lsp_gravity')):
        nodes,edges=bake_combined_plan(got)
        assert [(n['label'],n['x'],n['y']) for n in nodes]==[
            (n['label'],n['x'],n['y']) for n in got['combined'].nodes]
        assert edges==[list(map(float,e)) for e in got['combined'].machine_room_plan_edges or []]


def test_plan_sdf_stands_riser_while_kfp_and_iso_stay_byte_identical(tmp_path):
    import re
    from routes.module_f.emit import emit_merged
    got=merge_network(heads(),riser=system(True),mode='hsp_pump')
    iso=bake_combined_iso(got)[0]
    plain=emit_merged(got['combined'],tmp_path/'plain',iso_nodes=iso)
    shown=emit_merged(got['combined'],tmp_path/'shown',iso_nodes=iso,
                      plan_nodes=bake_combined_plan(got)[0])
    for key in ('kfp','kfp_iso','sdf_iso','slf','slf_iso'):
        assert Path(plain[key]).read_bytes()==Path(shown[key]).read_bytes(),key
    def read(path):
        text=Path(path).read_text(encoding='utf-8')
        pos={k:(float(x),float(y)) for k,x,y in re.findall(
            r'label="([^"]+)">\s*<Position x="([^"]+)" y="([^"]+)"',text)}
        return text,pos
    a,pa=read(plain['sdf']); b,pb=read(shown['sdf'])
    # 수리계산 값(배관 길이·rise·관경, 노드 표고·경계)은 한 글자도 안 바뀐다
    assert re.findall(r'<Pipe [^>]*>',a)==re.findall(r'<Pipe [^>]*>',b)
    assert re.findall(r'<Node [^>]*>',a)==re.findall(r'<Node [^>]*>',b)
    assert pa['n2']==pa['n3']                                  # 종전: 한 점으로 접힘
    assert pb['n2'][0]==pb['n3'][0] and pb['n3'][1]>pb['n2'][1]  # 이제 세로로 선다


def test_merge_preview_plan_view_uses_the_same_plan_function():
    from test_module_f_merge_view import _client,_sid
    c=_client(); sid,sess=_sid(c)
    sess['merged']=merge_network(heads(),riser=system(True),mode='hsp_pump')
    sess['supply_mode']='hsp_pump'
    plan=c.get(f'/api/module-f/merge/preview?sid={sid}').get_json()
    iso=c.get(f'/api/module-f/merge/preview?sid={sid}&iso=1').get_json()
    at={n['label']:(n['x'],n['y']) for n in plan['view']['nodes']}
    expect={n['label']:(n['x'],n['y']) for n in bake_combined_plan(sess['merged'])[0]}
    assert at==pytest.approx(expect)
    assert at['n3'][1]-at['n2'][1]==pytest.approx(2000)
    iso_at={n['label']:(n['x'],n['y']) for n in iso['view']['nodes']}
    assert iso_at==pytest.approx({n['label']:(n['x'],n['y']) for n in bake_combined_iso(sess['merged'])[0]})
