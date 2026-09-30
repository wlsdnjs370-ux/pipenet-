"""The real H preview uses the same scoped graph in plan and isometric modes."""
import time

from test_module_h_browser import h_ui
from routes.module_f import api_design, overrides


def test_scoped_plan_iso_and_tables_keep_local_cycle_only(h_ui, monkeypatch):
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    page, sess, _, tmp = h_ui
    monkeypatch.setattr(overrides, 'path_for', lambda key: str(tmp/'overrides.json'))
    monkeypatch.setattr(api_design, '_dia_texts', lambda _sess: [])
    def run(session, phase, function):
        result = function()
        session['job'] = dict(state='done', result=result, phase=phase, started=time.time(),
                              ended=time.time(), error=None)
    monkeypatch.setattr(api_design, '_run_job', run)
    points = [(0,0),(1000,0),(1000,1000),(0,1000),(2000,0),
              (3000,0),(3000,1000),(2000,1000),(3500,0),(3500,1000)]
    board = EditBoard('h-area-scope-browser', points,
        {(0,1),(1,2),(2,3),(0,3),(1,4),(4,5),(5,6),(6,7),(4,7),(5,8),(6,9)},
        [(3500,0,30),(3500,1000,30)])
    board.sources = [0]; board.valves = [0]; board.network_mode = 'loop'
    for disk in board.disks:
        board.set_head_kind(disk, '상향식')
    sess['edit'] = EditSession(board, key=board.key, out_dir=str(tmp))
    sess['design'] = None
    base = page.url.replace('/module-h','')
    result = page.request.post(base+'/api/module-f/edit/worst', data={
        'sid':sess['id'], 'selection_mode':'area_all', 'zones':[[1900,-100,3600,1100]]})
    assert result.ok, result.text()
    edit = result.json()['state']
    assert edit['worst']['area_scope']['cycle_rank'] == 1
    built = page.request.post(base+'/api/module-f/design/build', data={
        'sid':sess['id'], 'selection_mode':'area_all', 'review_default_mm':65})
    assert built.ok, built.text()
    assert sess['job']['result']['ok'], sess['job']['result']
    page.evaluate('''async ({sid,edit})=>{
      Object.assign(__mf,{sid,slot:'plan',method:'manual',zones:edit.worst.zones,
        world:{bounds:{minx:-500,miny:-500,maxx:4000,maxy:1500},bundles:[]}});
      __hTest.setEdit(edit);__hTest.setStage('design');await __hTest.preview();
    }''', dict(sid=sess['id'],edit=edit))
    page.wait_for_function('__mf.design?.tables?.nozzles?.length===2')
    page.wait_for_function("ModuleHAttributes.getState()?.records.length>0 && document.querySelector('#busy').classList.contains('hidden')")
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    assert page.evaluate('__mf.edit.worst.corridor.length') == 8
    assert sess['design']['got']['cycle_rank'] == 1
    assert set(sess['design']['got']['edge_ref'].values()) == sess['worst']['edges']
    page.screenshot(path=str(tmp/'area-scope-plan.png'))
    page.check('#ha-iso')
    page.wait_for_function("!document.querySelector('#dg-plan').checked")
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    page.screenshot(path=str(tmp/'area-scope-iso.png'))
    page.uncheck('#ha-iso')
    assert page.locator('#dg-plan').is_checked()
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
