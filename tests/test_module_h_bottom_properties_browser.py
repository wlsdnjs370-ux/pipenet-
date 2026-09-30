"""Table-only bottom dock retains drawing selection, evidence and live edits."""
import time

import ezdxf

from test_module_h_browser import h_ui, ROOT, menu
from routes.module_f import api_design, overrides


def prepare(h_ui, monkeypatch):
    monkeypatch.syspath_prepend(str(ROOT/'cad_project_editor_g'))
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    page,sess,_,tmp=h_ui
    monkeypatch.setattr(overrides,'path_for',lambda key:str(tmp/'overrides.json'))
    def run(session,phase,fn):
        result=fn();session['job']={'state':'done','result':result,'phase':phase,'started':time.time(),'ended':time.time(),'error':None}
    monkeypatch.setattr(api_design,'_run_job',run)
    texts=[(2500,100,65,0),(5100,2000,40,90),(10100,2000,40,90)]
    monkeypatch.setattr(api_design,'_dia_texts',lambda s:[t[:3] for t in texts])
    doc=ezdxf.new();doc.layers.new('FIRE-DIA')
    for x,y,d,angle in texts:
        doc.modelspace().add_text(f'{d}A',dxfattribs=dict(insert=(x,y),rotation=angle,height=120,layer='FIRE-DIA'))
    path=tmp/'source.dxf';doc.saveas(path);sess['dxf']=str(path)
    points=[(0,0),(5000,0),(10000,0),(15000,0),(5000,4000),(10000,4000)]
    board=EditBoard('h-bottom-properties',points,{(0,1),(1,2),(2,3),(1,4),(2,5)},[(5000,4000,30),(10000,4000,30)])
    board.sources=[0];board.valves=[0]
    for disk in board.disks:board.set_head_kind(disk,'상향식')
    sess['edit']=EditSession(board,key=board.key,out_dir=str(tmp));sess['design']=None
    base=page.url.replace('/module-h','')
    response=page.request.post(base+'/api/module-f/edit/worst',data={'sid':sess['id'],'selection_mode':'area_all','zones':[[4500,3500,10500,4500]]})
    assert response.ok,response.text()
    edit=response.json()['state']
    world={'bounds':{'minx':-1000,'miny':-1000,'maxx':16000,'maxy':6000},'bundles':[]}
    page.evaluate('''({sid,edit,world})=>{Object.assign(__mf,{sid,world,slot:'plan',method:'manual',zones:edit.worst.zones});__hTest.setEdit(edit);__hTest.setStage('design');}''',dict(sid=sess['id'],edit=edit,world=world))
    page.wait_for_function("ModuleHAttributes.getState()?.annotations.length===3 && document.querySelector('#busy').classList.contains('hidden')")
    # Keep fixture click points above the bottom dock as canvas height changes.
    page.evaluate("__mf.view={scale:.06,ox:-1000,oy:2000-(document.querySelector('#cv').getBoundingClientRect().height-422)/.06};__hTest.draw()")
    return page,sess,tmp


def test_click_pipe_table_only_dock_live_edit_and_undo(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    pos=page.evaluate('''()=>{const b=document.querySelector('#cv').getBoundingClientRect();return [b.left+__mf.toScreenX(5000),b.top+__mf.toScreenY(2000)]}''')
    page.mouse.click(*pos)
    assert page.evaluate('__mf.design.sel'),page.evaluate('''()=>({mode:__mf.edit.mode,plan:document.querySelector('#dg-plan').checked,ne:document.querySelector('#ne-panel').getClientRects().length,drawer:document.body.classList.contains('h-drawer-open'),records:ModuleHAttributes.getState().records.map(r=>({label:r.label,kind:r.kind,xy:r.xy})),view:__mf.view})''')
    assert page.locator('#h-properties').is_hidden()
    menu(page,'#h-table')
    page.wait_for_selector('#h-properties:not([hidden]) #hp-rows .hp-selected')
    assert page.locator('#dg-ins').is_hidden()
    assert page.locator('#hp-reference, #hp-source-canvas, #hp-source-details').count()==0
    dock=page.locator('#h-properties').bounding_box()
    table=page.locator('#h-properties .hp-scroll').bounding_box()
    assert abs(table['y']+table['height']-(dock['y']+dock['height']-1))<2
    assert table['height']>dock['height']*.65  # Freed space belongs to table rows.
    selected=page.locator('#hp-rows .hp-selected').get_attribute('data-label')
    row=page.locator(f'#hp-rows tr[data-label="{selected}"]')
    assert row.locator('[data-field=dn]').input_value()=='40'
    before=[dict(r) for r in sess['design']['tables'].pipes]
    row.locator('[data-field=dn]').select_option('80')
    page.wait_for_function("label=>ModuleHProperties.getData()?.pipes.find(r=>r.label===label)?.dia===80",arg=selected)
    result=next(r for r in sess['design']['tables'].pipes if r['label']==selected)
    assert result['dia']==80 and result['inner_mm']>80
    assert result['bore_provenance']['text_raw']=='40A'
    assert result['bore_provenance']['text_layer']=='FIRE-DIA'
    assert result['bore_provenance']['manual']['value_mm']==80
    assert page.locator('#hp-reference').count()==0
    # Focused number entry must not disappear on selection-only polling refresh.
    row.locator('[data-field=c]').fill('135')
    page.wait_for_timeout(350)
    assert row.locator('[data-field=c]').input_value()=='135'
    row.locator('[data-field=c]').press('Tab')
    page.wait_for_function("label=>ModuleHProperties.getData()?.pipes.find(r=>r.label===label)?.c===135",arg=selected)
    page.click('#hp-undo')
    page.wait_for_function("label=>ModuleHProperties.getData()?.pipes.find(r=>r.label===label)?.c===120",arg=selected)
    page.click('#hp-undo')
    page.wait_for_function("label=>ModuleHProperties.getData()?.pipes.find(r=>r.label===label)?.dia===40",arg=selected)
    page.wait_for_function("label=>ModuleHAttributes.getState()?.records.find(r=>r.label===label)?.nominal_mm===40",arg=selected)
    assert [(r['label'],r['length']) for r in sess['design']['tables'].pipes]==[(r['label'],r['length']) for r in before]
    # Table -> canvas uses the same selected record, not a separate graph edit.
    other=next(r['label'] for r in before if r['label']!=selected)
    page.locator(f'#hp-rows tr[data-label="{other}"] td').first.click()
    page.wait_for_function("label=>__mf.design.sel?.label===label",arg=other)
    page.locator(f'#hp-rows tr[data-label="{selected}"] td').first.click()
    assert page.evaluate('__mf.design.sel.label')==selected
    page.screenshot(path=str(tmp/'bottom-dock.png'))
    # Collapse keeps the selection and frees drawing space.
    page.click('#hp-fold');assert page.locator('#h-properties').bounding_box()['height']<=50
    page.click('#hp-fold');assert page.locator('#hp-rows').is_visible()


def test_small_table_only_dock_scrolls(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    page.evaluate("ModuleHProperties.open()")
    page.wait_for_selector('#hp-rows tr')
    page.set_viewport_size(dict(width=390,height=760))
    page.locator('#hp-rows tr td').first.click()
    box=page.locator('#h-properties').bounding_box()
    assert box['x']>=0 and box['x']+box['width']<=390
    assert page.locator('#h-properties .hp-scroll').evaluate('n=>n.scrollWidth>n.clientWidth')
    assert page.locator('#hp-reference, #hp-source-canvas, #hp-source-details').count()==0
    page.screenshot(path=str(tmp/'bottom-dock-narrow.png'))


def test_plan_head_click_selects_its_canonical_node_row(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    head=page.evaluate("ModuleHAttributes.getState().records.find(r=>r.kind==='head')")
    assert head['node_label']
    pos=page.evaluate('''xy=>{const b=document.querySelector('#cv').getBoundingClientRect();return [b.left+__mf.toScreenX(xy[0]),b.top+__mf.toScreenY(xy[1])]}''',head['xy'][0])
    page.mouse.click(*pos)
    assert page.locator('#h-properties').is_hidden()
    menu(page,'#h-table')
    page.wait_for_function("label=>document.querySelector('#hp-kind').value==='node' && document.querySelector('#hp-rows .hp-selected')?.dataset.label===label",arg=head['node_label'])
    assert page.locator('#hp-rows .hp-selected [data-field=k]').is_enabled()
