"""H review export, compact UI, table transactions and selective DXF regression."""
from copy import deepcopy
import xml.etree.ElementTree as ET

import pytest

from test_module_h_browser import h_ui, seed, settle
from test_module_f_sizing import tables


def test_annotation_reader_preserves_native_nested_transform_and_drops_geometry(tmp_path, monkeypatch):
    import ezdxf
    from src.pipenet_converter.dxf import annotation_document as reader
    from src.pipenet_converter.dxf.diameter_annotations import enrich_diameter_annotations
    doc=ezdxf.new('R2013');doc.layers.new('FIRE')
    child=doc.blocks.new('child',base_point=(10,20))
    child.add_text('65',dxfattribs={'insert':(20,40),'rotation':90,'height':3})
    parent=doc.blocks.new('parent')
    parent.add_blockref('child',(100,200),dxfattribs={'rotation':30,'xscale':-2,'yscale':2})
    ins=doc.modelspace().add_blockref('parent',(1000,2000),dxfattribs={'rotation':45,'layer':'FIRE'})
    ins.add_attrib('NAME','ignored',(100,200))
    doc.modelspace().add_polyline2d([(0,0),(0,100),(100,100)])
    doc.modelspace().add_mtext('100',dxfattribs={'insert':(3000,4000),'rotation':90})
    doc.modelspace().add_text('50',dxfattribs={'insert':(5000,6000),'extrusion':(0,0,-1),'rotation':30})
    for i in range(2000):doc.modelspace().add_line((i,0),(i,200))
    path=tmp_path/'text.dxf';doc.saveas(path)
    filtered=reader.read_annotation_document(path)
    assert len(filtered.modelspace().query('LINE'))==0
    assert len(filtered.modelspace().query('INSERT'))==1
    def text_points(entities):
        for e in entities:
            if e.dxftype()=='INSERT':yield from text_points(e.virtual_entities())
            elif e.dxftype() in ('TEXT','MTEXT'):
                p=e.ocs().to_wcs(e.dxf.insert) if e.dxftype()=='TEXT' else e.dxf.insert
                yield (p.x,p.y,int(e.plain_text()))
    points=list(text_points(ezdxf.readfile(path).modelspace()))
    fast=enrich_diameter_annotations(path,points)
    monkeypatch.setattr(reader,'read_annotation_document',lambda path,**_kwargs:ezdxf.readfile(path))
    original=enrich_diameter_annotations(path,points)
    assert fast==original
    assert all(r.rotation_deg is not None for r in fast)


def test_unknown_plan_review_sdf_retains_graph_and_source_values(tmp_path):
    tbl=tables()
    from services.cad_import.design.emit import emit_design_sdf
    tbl.pipes[0].update(dia=0,bore_provenance={'policy':'drawing_first_v1','block_export':True})
    tbl.pipes[1].update(bore_provenance={'policy':'drawing_first_v1','block_export':True})
    before=deepcopy(tbl)
    with pytest.raises(ValueError):emit_design_sdf(tbl,tmp_path/'normal.sdf')
    out=emit_design_sdf(tbl,tmp_path/'review.sdf',plan_review=True)
    root=ET.parse(out).getroot();pipes={p.get('label'):p for p in root.iter('Pipe')}
    assert len(pipes)==len(tbl.pipes)
    assert float(pipes['1'].get('bore'))==0
    assert float(pipes['2'].get('bore'))==0
    assert float(pipes['3'].get('bore'))==.025
    assert 'NOT CALCULATION READY' in root.findtext('.//Network-spray/Title')
    assert out.with_suffix('.review.json').is_file()
    assert tbl==before


def test_h_plan_export_adapter_allows_unknown_bore_without_relaxing_f(tmp_path,monkeypatch):
    from routes.module_f import api_design
    tbl=tables();tbl.pipes[0].update(dia=0,bore_provenance={'policy':'drawing_first_v1','block_export':True})
    sess={'id':'h-draft-export','key':'review','design':{'tables':tbl,'got':{}},
          'design_settings':dict(api_design._DEFAULT_SETTINGS,diameter_policy='drawing_first_v1')}
    monkeypatch.setattr(api_design,'_design_stale',lambda s:None)
    path,error=api_design.emit_design_files(sess,tmp_path)
    assert path and not error,error
    assert float(next(p for p in ET.parse(path).getroot().iter('Pipe') if p.get('label')=='1').get('bore'))==0
    assert 'NOT CALCULATION READY' in path.read_text(encoding='utf-8')
    assert sess['design_review_path'].endswith('.review.json')
    sess['design_settings']['diameter_policy']='legacy'
    path,error=api_design.emit_design_files(sess,tmp_path)
    assert not path and '관경' in error


@pytest.mark.parametrize('kind,field,value',[('nodes','x',float('nan')),('pipes','c',0),('pipes','dia',-1),('nozzles','in','missing')])
def test_plan_review_does_not_relax_graph_or_numeric_validation(kind,field,value):
    from src.pipenet_converter.sdf.plan_review import review_tables
    tbl=tables();getattr(tbl,kind)[0][field]=value
    with pytest.raises(ValueError):review_tables(tbl)


def test_h_clean_controls_start_unarmed(h_ui):
    page,sess,_,_=h_ui
    assert not page.locator('#ed-zone-arm').is_checked()
    seed(page,sess,'edit');page.click('#h-more');page.click('#h-task');settle(page)
    assert page.locator('#h-region-draw').is_visible()
    assert page.locator('.emode[data-mode="원클릭"]').inner_text()=='시작노드 정의'
    page.click('#h-region-draw');assert page.locator('#ed-zone-arm').is_checked()
    page.click('.emode[data-mode="원클릭"]');assert not page.locator('#ed-zone-arm').is_checked()
    assert page.locator('#open-zoom-window').is_hidden()
    assert page.locator('#ed-network-note').is_hidden()
    assert page.locator('#dg-review-notice').is_hidden()
    assert page.locator('#dg-build').inner_text()=='배관 속성 확정'


def test_integrated_table_bidirectional_edit_and_undo(h_ui):
    from test_module_f_sizing import merged
    from routes.module_f import network_edit as ne
    page,sess,_,_=h_ui
    sess['merged']=merged(tables());sess['merge_summary']={'merged':True}
    seed(page,sess,'merge')
    page.evaluate("async()=>{await __hTest.merged();await moduleFNetworkEditor.refresh();}")
    page.wait_for_selector('#h-properties')
    page.wait_for_selector('#hp-rows tr[data-label="1"]')
    page.locator('#hp-rows tr[data-label="1"]').click()
    assert page.evaluate('__mf.mergeSel')=={'kind':'pipe','label':'1'}
    page.evaluate("__hTest.select('mgpipe','2')")
    page.wait_for_selector('#hp-rows tr[data-label="2"].hp-selected')
    c=page.locator('#hp-rows tr[data-label="2"] input[data-field="c"]')
    c.fill('130');c.press('Tab')
    page.wait_for_function("document.querySelector('#hp-message').textContent==='저장됨'")
    assert next(r for r in sess['merged']['combined'].pipes if r['label']=='2')['c']==130
    page.click('#hp-undo')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && document.querySelector('#hp-rows tr[data-label=\"2\"] input[data-field=c]').value==='120'")
    assert next(r for r in sess['merged']['combined'].pipes if r['label']=='2')['c']==120
    page.evaluate("__hTest.select('mgnode','a')")
    page.wait_for_selector('#hp-rows tr[data-label="a"].hp-selected')
    assert page.locator('#hp-kind').input_value()=='node'
    assert page.locator('#ne-toolbar').is_hidden()
    assert page.locator('#dg-ins').is_hidden()
    # Table movement does not mutate calculation coordinates.
    before=deepcopy(sess['merged']['combined'].nodes)
    head=page.locator('#h-properties header');box=head.bounding_box()
    page.mouse.move(box['x']+100,box['y']+15);page.mouse.down();page.mouse.move(box['x']+50,box['y']-85);page.mouse.up()
    assert before==sess['merged']['combined'].nodes
    page.evaluate("__hTest.select('mgnode','h')")
    page.wait_for_selector('#hp-rows tr[data-label="h"].hp-selected')
    k=page.locator('#hp-rows tr[data-label="h"] input[data-field=k]')
    k.fill('90');k.press('Tab')
    page.wait_for_function("document.querySelector('#hp-message').textContent==='저장됨' && document.querySelector('#hp-rows tr[data-label=h] input[data-field=k]').value==='90'")
    pressure=page.locator('#hp-rows tr[data-label="h"] input[data-field=pressure]')
    pressure.fill('1.5');pressure.press('Tab')
    page.wait_for_function("document.querySelector('#hp-message').textContent==='저장됨' && document.querySelector('#hp-rows tr[data-label=h] input[data-field=pressure]').value==='1.5'")
    head=sess['merged']['combined'].nozzles[0]
    assert head['k_factor_si']==90 and head['required_pressure_bar']==1.5
    from test_module_h_browser import ROOT
    page.screenshot(path=str(ROOT/'data/module_h_diameter_review/integrated-properties.png'))


def test_prepared_world_is_request_local_and_checks_source_and_crop(tmp_path):
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff,stage1
    path=tmp_path/'drawing.dxf';path.write_text('original')
    world=stage1.World();world._source_signature=handoff.source_signature(path)
    world._work_regions=[[0,0,100,100]];world.texts=[('mapped',7,1,2,3,'100')]
    with handoff.using_world('drawing',world):
        assert handoff.prepared_world('other',path,world._work_regions) is None
        assert handoff.prepared_world('drawing',path,[]) is None
        got=handoff.prepared_world('drawing',path,world._work_regions)
        assert got is not world and got.texts is world.texts
        path.write_text('different source')
        assert handoff.prepared_world('drawing',path,world._work_regions) is None
    assert handoff.prepared_world('drawing',path,world._work_regions) is None


def test_read_observer_does_not_cause_false_extraction_failure():
    import threading,time
    from routes.module_f import cancellation as c
    sess={};token=c.Operation('reader-handoff-test');token.reserve()
    with c._LOCK:c._SESSION_READERS[id(sess)]=1
    outcome=[]
    worker=threading.Thread(target=lambda:(token.claim(sess,wait_readers=1),outcome.append('acquired')))
    worker.start();time.sleep(.05)
    assert not outcome
    with c._LOCK:
        c._SESSION_READERS.pop(id(sess));c._READ_CONDITION.notify_all()
    worker.join(2)
    assert outcome==['acquired']
    token.release()


def test_diameter_progress_gap_recovers_representative_snapshot():
    from routes.module_h_progress import PreviewStore
    from src.pipenet_converter.progress import ProgressEvent
    store=PreviewStore()
    for i in range(240):
        store.publish('test',ProgressEvent('diameter','대응',{'edge':[i,i+1],'xy':[[i,0],[i+1,0]],'evidence':{'text_mm':100}}))
    got=store.read('test',0)
    assert got['gap'] and len(got['diameter_snapshot'])==240
    assert got['snapshot_cursor']==240


@pytest.mark.parametrize('slot',['system','machineroom'])
def test_extract_waits_for_point_preview_without_repeated_clicks(h_ui,slot):
    page,sess,_,_=h_ui
    seed(page,sess,'sub')
    result=page.evaluate('''async slot=>{
      __mf.slot=slot;__mf.sub={picks:[[0,0],[1000,0]],arm:null};
      const original=window.fetch;let graphDone=false,extracts=0,premature=false;
      window.fetch=async(url,options)=>{
        const path=String(url);
        if(path.includes('/sub/graph')){
          await new Promise(resolve=>setTimeout(resolve,250));graphDone=true;
          return new Response(JSON.stringify({ok:true,nodes:[[0,0],[1000,0]],edges:[[0,1,1000,false]],layers:[],chosen:[],forced_penalty_mm:2500}));
        }
        if(path.includes('/'+slot+'/extract')){
          extracts++;premature=!graphDone;
          return new Response(JSON.stringify(premature?{ok:false,error:'writer busy'}:{ok:true,summary:{nodes:2,pipes:1,total_m:1},fixed:{}}),{status:premature?409:200});
        }
        return original(url,options);
      };
      try{__hTest.loadSubGraph();await __hTest.subExtract(false);return {extracts,premature,summary:__mf.sub.summary};}
      finally{window.fetch=original;}
    }''',slot)
    assert result==dict(extracts=1,premature=False,summary=dict(nodes=2,pipes=1,total_m=1))


def test_table_custom_head_properties_survive_later_flow_edit():
    from core.network_editor import Network,apply_edit,EditError
    from test_module_f_network_editor import table
    from routes.module_f import network_edit as ne
    original=Network.from_tables(table());before=deepcopy(original.tables)
    changed,_=apply_edit(original,dict(op='head',target='3',nozzle='SP-HEAD',flow=90,k_factor_si=90,required_pressure_bar=1.5),ne.catalog())
    changed,_=apply_edit(changed,dict(op='head',target='3',nozzle='SP-HEAD',flow=100),ne.catalog())
    head=changed.tables.nozzles[0]
    assert head['label']=='H1' and head['out']=='@/3'
    assert head['k_factor_si']==90 and head['required_pressure_bar']==1.5
    assert head['flow_lmin']==100 and head['definition_policy']=='drawing_first_v1'
    assert original.tables==before
    with pytest.raises(EditError):
        apply_edit(original,dict(op='head',target='3',nozzle='SP-HEAD',flow=90,k_factor_si='NaN'),ne.catalog())


def test_edit_existing_head_retains_status_and_does_not_add_a_junction_head():
    from core.network_editor import Network,apply_edit,EditError
    from routes.module_f import network_edit as ne
    tbl=tables();tbl.nozzles[0]['status']='0'
    changed,_=apply_edit(Network.from_tables(tbl),dict(op='head',target='h',nozzle='SP-HEAD',flow=80,k_factor_si=90),ne.catalog())
    assert changed.tables.nozzles[0]['status']=='0'
    assert changed.tables.nozzles[0]['label']=='H1'
    with pytest.raises(EditError):
        apply_edit(Network.from_tables(tbl),dict(op='head',target='a',nozzle='SP-HEAD',flow=80),ne.catalog())
