"""Adapter, isolated HTTP, stale-proposal and SDF loss/boundary regression."""
from copy import deepcopy
import json
import xml.etree.ElementTree as ET
from pathlib import Path
import sys

import pytest
from flask import Flask

from routes.module_f import api_sizing, jobs
from routes.module_f.common import _boot
from routes.module_f.hydraulic_sizing import fingerprint, prepare
from src.pipenet_converter.hydraulics.sizing import propose

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'core'))


def tables():
    _boot()
    from core.remote30_full_network import CombinedTables
    return CombinedTables(
        nodes=[dict(label=n,x=x,y=y,elevation=0,io_node='Input' if n=='s' else 'No')
               for n,x,y in [('s',0,0),('a',10000,10000),('b',10000,-10000),('h',20000,0)]],
        pipes=[dict(label=str(i),**{'in':a,'out':b},type='KSD 3507',dia=25,
                    length=20,elev=0,c=120,status='Normal',group='Unset')
               for i,(a,b) in enumerate([('s','a'),('a','h'),('s','b'),('b','h'),('a','b')],1)],
        nozzles=[dict(label='H1',**{'in':'h','out':'@/1'},status='1',lib='SP-HEAD',flow_lmin=80,flow_m3s=80/60000)],
        fittings=[dict(pipe='1',**{'in':'s','out':'a'},type='elbow',count='1')],
        equipment=[],pumps=[],valves=[],meta=[],machine_room_plan_edges=[],layout_status={})


def merged(tbl):
    return dict(combined=tbl,parts={'plan':[str(n['label']) for n in tbl.nodes],
                'system':[],'machineroom':[]},mode='hsp_pump')


def test_adapter_real_inner_and_no_fitting_double_count():
    tbl=tables();tbl.pipes[0]['eq_len']=999  # cached sum must NOT be counted again
    net,cfg,warnings=prepare(tbl,{})
    assert net.pipes[0].sizes[0].inner_mm != 25
    assert 0 < net.pipes[0].sizes[0].equivalent_m < 10
    # The authoritative SLF is K80 in kgf/cm2 units, not exactly K80 per bar.
    assert net.nozzles[0].k_lpm_sqrt_bar==pytest.approx(80.7848,rel=1e-5)
    assert all(p.role=='unknown' for p in net.pipes)
    assert any('역할 미지정' in w for w in warnings)


def test_unresolved_equipment_needs_explicit_loss_and_locks():
    tbl=tables();tbl.equipment=[dict(pipe='1',desc='A/V',eq_len=0)]
    with pytest.raises(ValueError,match='미확정'):
        prepare(tbl,{})
    net,_,_=prepare(tbl,{'pipes':{'1':{'equivalent_m':9.5}}})
    assert net.pipes[0].locked and len(net.pipes[0].sizes)==1
    assert net.pipes[0].sizes[0].equivalent_m==9.5


def test_rejects_unknown_diameter_and_no_active_heads():
    tbl=tables();tbl.pipes[0]['dia']=29
    with pytest.raises(ValueError,match='실제 내경'):
        prepare(tbl,{})
    with pytest.raises(ValueError,match='작동 노즐'):
        prepare(tables(),{'active_nozzles':[]})
    tbl=tables();tbl.meta=[('부속 판정 불가','1')]
    with pytest.raises(ValueError,match='부속 판정 미확정'):
        prepare(tbl,{})


def test_h_unknown_sizing_scope_locks_drawing_and_never_changes_input():
    tbl=tables();tbl.pipes[0].update(dia=0,bore_provenance={'policy':'drawing_first_v1','source':'unresolved','block_export':True})
    before=deepcopy(tbl)
    net,_,_=prepare(tbl,{})
    assert not net.pipes[0].locked and len(net.pipes[0].sizes)>1
    assert all(p.locked for p in net.pipes[1:])
    net,_,_=prepare(tbl,{'diameter_scope':'all'})
    assert not any(p.locked for p in net.pipes)
    assert tbl.pipes==before.pipes
    with pytest.raises(ValueError,match='잠글'):
        prepare(tbl,{'pipes':{'1':{'locked':True}}})


@pytest.fixture
def client(tmp_path,monkeypatch):
    sess=jobs._new_session(key='sizing-test',merged=merged(tables()))
    def run(sess,phase,fn):
        result=fn();sess['job']={'state':'done','result':result}
    monkeypatch.setattr(api_sizing,'_run_job',run)
    app=Flask(__name__);app.testing=True
    api_sizing.register(app,UPLOAD_DIR=tmp_path)
    yield app.test_client(),sess
    jobs._SESSIONS.pop(sess['id'],None)


def run(c,s):
    state=c.get('/api/module-f/merge/sizing/state',query_string={'sid':s['id']}).json
    r=c.post('/api/module-f/merge/sizing/run',json={'sid':s['id'],
        'fingerprint':state['basis']['fingerprint'],'options':{'mode':'joint'}})
    assert r.status_code==200,r.json


def test_api_proposal_does_not_modify_source_and_stale_is_blocked(client):
    c,s=client;before=fingerprint(s['merged']['combined'])
    run(c,s)
    report=s['sizing_report']
    assert report['feasible'],report['violations']
    assert fingerprint(s['merged']['combined'])==before
    assert c.get('/api/module-f/merge/sizing/download',query_string={'sid':s['id']}).status_code==200
    s['merged']['combined'].pipes[0]['length']+=1
    assert c.get('/api/module-f/merge/sizing/state',query_string={'sid':s['id']}).json['stale']
    assert c.get('/api/module-f/merge/sizing/download',query_string={'sid':s['id']}).status_code==409
    assert c.post('/api/module-f/merge/sizing/emit',json={'sid':s['id'],'proposal_id':report['proposal_id']}).status_code==409


@pytest.mark.parametrize('manual',[False,True])
def test_review_sdf_keeps_cycles_and_explicit_pressure_and_total_losses(client,manual):
    c,s=client
    if manual:
        from core.network_editor import Network, apply_edit
        from routes.module_f.network_edit import catalog
        tbl = s['merged']['combined']
        tbl.fittings = []
        tbl.meta = [('부속 판정 불가','1')]
        tbl.unresolved = {'kind_items':[dict(pipe_label='1',node_label='a',n=1)]}
        edited,_ = apply_edit(Network.from_tables(tbl),dict(op='fitting',target='1',
            node='a',fitting='ELBOW_90_STD',count=1),catalog())
        s['merged']['combined'] = edited.to_tables()
    before=fingerprint(s['merged']['combined']);run(c,s)
    report=s['sizing_report']
    r=c.post('/api/module-f/merge/sizing/emit',json={'sid':s['id'],'proposal_id':report['proposal_id']})
    assert r.status_code==200,r.json
    for kind in ('sdf','sdf_iso'):
        root=ET.parse(s['sizing_files'][kind]).getroot()
        assert len(list(root.iter('Pipe')))==5 and len(list(root.iter('Nozzle')))==1
        assert not list(root.iter('Pump-fan'))
        spec=root.find('.//Node[@label="s"]/Calculation-spec')
        assert float(spec.get('pressure'))==pytest.approx(101325+report['solution']['source_pressure_bar']*100000,abs=1)
        assert root.find('.//Design-options').get('specification-type')=='user-defined'
        assert not list(root.iter('Fitting')) # accounted exactly once as equipment
        assert len(list(root.iter('Equipment')))==1
        assert 'NOT VALIDATED' in root.find('.//Network-spray/Title').text
    assert fingerprint(s['merged']['combined'])==before
    exported=c.get('/api/module-f/merge/sizing/download',query_string={'sid':s['id'],'what':'report'})
    assert exported.status_code==200
    assert json.loads(exported.data)['export_compaction']['nodes_before']==4


def test_unsupported_pump_suction_and_control_valve_are_not_silently_dropped():
    tbl=tables();tbl.valves=[dict(label='PRV1')]
    with pytest.raises(ValueError,match='제어밸브'):
        prepare(tbl,{})
    tbl=tables();tbl.pumps=[{'in':'s','out':'a'}]
    with pytest.raises(ValueError,match='흡입측'):
        prepare(tbl,{})


def test_merged_native_valve_names_use_library_loss_without_skipping():
    from routes.module_f.network_edit import catalog
    tbl = tables()
    tbl.pipes[0]['dia'] = 100
    tbl.fittings = [dict(pipe='1',type=kind,count=1) for kind in ('gate','check','butterfly')]
    net,_,_ = prepare(tbl,{})
    wanted = {'VALVE_GATE','VALVE_SWING_CHECK','VALVE_BUTTERFLY'}
    expected = sum(r['lengths']['100'] for r in catalog()['fittings'] if r['id'] in wanted)
    assert net.pipes[0].sizes[0].equivalent_m == pytest.approx(expected)
    tbl.pipes[0]['dia'] = 25
    with pytest.raises(ValueError,match='gate 25A 등가길이 없음'):
        prepare(tbl,{})  # unsupported loss is still blocked, never zeroed


def test_stale_manual_equipment_loss_must_be_confirmed():
    tbl = tables()
    tbl.equipment = [dict(label='EF1',pipe='1',editor_node='a',
        editor_library='ELBOW_90_STD',count=1,eq_len=99)]
    with pytest.raises(ValueError,match='저장된 손실과 라이브러리'):
        prepare(tbl,{})
    net,_,_ = prepare(tbl,{'pipes':{'1':{'equivalent_m':8}}})
    assert net.pipes[0].locked and net.pipes[0].sizes[0].equivalent_m == 8


def test_real_background_job_with_cancellation_middleware_copies_only_outputs(tmp_path):
    import time
    from routes.module_f import cancellation
    class NoCopy:
        def __deepcopy__(self,memo):
            raise AssertionError('large CAD session must not be copied')
    sess=jobs._new_session(key='sizing-real-worker',merged=merged(tables()),world=NoCopy())
    original=fingerprint(sess['merged']['combined'])
    app=Flask(__name__);app.testing=True
    api_sizing.register(app,UPLOAD_DIR=tmp_path)
    cancellation.install(app)
    try:
        c=app.test_client()
        reply=c.post('/api/module-f/merge/sizing/run',json={'sid':sess['id'],
            'fingerprint':original,'options':{'mode':'pump_only'}})
        assert reply.status_code==200,reply.json
        deadline=time.monotonic()+10
        while sess['job']['state']=='run' and time.monotonic()<deadline:
            time.sleep(.01)
        assert sess['job']['state']=='done',sess['job']
        assert sess['sizing_report']['feasible']
        assert fingerprint(sess['merged']['combined'])==original
        op=cancellation.operation(sess['job']['operation'])
        assert not op.workspaces
    finally:
        jobs._SESSIONS.pop(sess['id'],None)
