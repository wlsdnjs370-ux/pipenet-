"""Live H attributes UI against an isolated real graph and real HTTP adapters."""
import time
import ezdxf
from test_module_h_browser import h_ui, ROOT
from routes.module_f import api_design, overrides


def test_automatic_attributes_shift_selection_batch_copy_and_live_overlay(h_ui,monkeypatch):
    monkeypatch.syspath_prepend(str(ROOT/'cad_project_editor_g'))
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    page,sess,requests,tmp=h_ui
    monkeypatch.setattr(overrides,'path_for',lambda key:str(tmp/'overrides.json'))
    def run(session,phase,fn):
        result=fn();session['job']={'state':'done','result':result,'phase':phase,'started':time.time(),'ended':time.time(),'error':None}
    monkeypatch.setattr(api_design,'_run_job',run)
    monkeypatch.setattr(api_design,'_dia_texts',lambda s:[(2500,100,65),(5100,2000,40),(10100,2000,40)])
    # Automatic H sources require real native text identity, not only XY/DN tuples.
    doc=ezdxf.new()
    for x,y,d,angle in [(2500,100,65,0),(5100,2000,40,90),(10100,2000,40,90)]:
        doc.modelspace().add_text(f'{d}A',dxfattribs=dict(insert=(x,y),height=120,rotation=angle))
    source=tmp/'attributes.dxf';doc.saveas(source);sess['dxf']=str(source)
    points=[(0,0),(5000,0),(10000,0),(15000,0),(5000,4000),(10000,4000)]
    board=EditBoard('h-diameter-browser',points,{(0,1),(1,2),(2,3),(1,4),(2,5)},[(5000,4000,30),(10000,4000,30)])
    board.sources=[0];board.valves=[0]
    for disk in board.disks:board.set_head_kind(disk,'상향식')
    sess['edit']=EditSession(board,key=board.key,out_dir=str(tmp));sess['design']=None
    base=page.url.replace('/module-h','')
    result=page.request.post(base+'/api/module-f/edit/worst',data={'sid':sess['id'],'selection_mode':'area_all','zones':[[4500,3500,10500,4500]]})
    assert result.ok,result.text()
    edit=result.json()['state']
    world={'bounds':{'minx':-1000,'miny':-1000,'maxx':16000,'maxy':6000},'bundles':[]}
    page.evaluate('''({sid,edit,world})=>{Object.assign(__mf,{sid,world,slot:'plan',method:'manual',zones:edit.worst.zones});__hTest.setEdit(edit);__hTest.setStage('design');}''',{'sid':sess['id'],'edit':edit,'world':world})
    page.wait_for_function("__mf.design?.settings?.diameter_policy==='drawing_first_v1' && document.querySelector('#busy').classList.contains('hidden')")
    page.wait_for_function("document.querySelector('#ha-annotation').options.length===4")
    page.wait_for_selector('#h-evidence:not([hidden]) #he-rows tr')
    assert '전체 6' in page.locator('#he-status').inner_text()
    page.locator('#he-rows tr').first.click()
    assert page.locator('#he-targets button').count()>0
    assert page.locator('#h-evidence-canvas').is_visible()
    assert ('POST','/api/module-f/design/build') in requests
    assert page.locator('#dg-review-dia').is_hidden()
    # LAN HTTP has no crypto.randomUUID; edits must still have operation IDs.
    page.evaluate("Object.defineProperty(crypto,'randomUUID',{value:undefined,configurable:true})")
    page.click('#ha-open');page.click('[data-mode=click]')
    # Keep the test graph left of the floating panel, using the native camera.
    page.evaluate("__mf.view={scale:.05,ox:-1000,oy:-2000};__hTest.draw()")
    def click_xy(x,y,shift=False):
        pos=page.evaluate('''([x,y])=>{const r=document.querySelector('#cv').getBoundingClientRect();return{x:r.left+__mf.toScreenX(x),y:r.top+__mf.toScreenY(y)}}''',[x,y])
        page.mouse.click(pos['x'],pos['y']) if not shift else None
        if shift:
            page.keyboard.down('Shift');page.mouse.click(pos['x'],pos['y']);page.keyboard.up('Shift')
    click_xy(2500,0);click_xy(7500,0,True)
    assert page.locator('#ha-selected').inner_text().startswith('2 배관')
    page.click('#ha-manual summary');page.select_option('#ha-diameter','80');page.click('#ha-apply')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && __mf.design.tables.pipes.filter(r=>r.dia===80).length>=2")
    page.wait_for_function("!document.querySelector('#ha-undo').disabled")
    page.click('#ha-undo')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && !__mf.design.tables.pipes.some(r=>r.dia===80)")
    # Clipboard mapping uses actual selected stable keys, never OS clipboard text.
    page.click('#ha-clear');click_xy(5000,2000);page.click('#ha-copy')
    click_xy(10000,2000);page.click('#ha-paste')
    page.wait_for_selector('#ha-copy-rows select')
    assert page.locator('#ha-copy-preview').is_visible()
    # Observable progress keeps the same kind, phase and actual processed geometry.
    page.evaluate("window.dispatchEvent(new CustomEvent('module-h-diameter-progress',{detail:{kind:'diameter',phase:'반복 구조 대응',data:{edge:[1,4],xy:[[5000,0],[5000,4000]],evidence:{text_mm:40,definition_source:'drawing_repeat'}}}}))")
    assert '반복 구조 대응' in page.locator('#ha-count').inner_text()
    output=ROOT/'data/module_h_diameter_review';output.mkdir(exist_ok=True)
    page.screenshot(path=str(output/'attributes-workbench.png'))
