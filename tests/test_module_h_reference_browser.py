"""Same native context menu and source-text picking in plan and iso."""
import json

import pytest

from test_module_h_browser import h_ui, settle
from test_module_h_bottom_properties_browser import prepare
from test_module_h_view_evidence_browser import capture_canvas


def screen(page, xy, project=False):
    return page.evaluate('''({xy,project})=>{const p=project?ModuleHView.projectPoint(xy):xy,b=document.querySelector('#cv').getBoundingClientRect();
      return [b.left+__mf.toScreenX(p[0]),b.top+__mf.toScreenY(p[1])];}''', dict(xy=xy, project=project))


@pytest.mark.parametrize('iso', [False, True])
def test_reference_menu_pick_commit_and_ctrl_z_in_both_views(h_ui, monkeypatch, iso):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    if iso:
        page.check('#ha-iso');settle(page)
    page.uncheck('#ha-show')
    record = page.evaluate("ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.nominal_mm===40&&!r.generated)")
    point = page.evaluate('''r=>{const seg=ModuleHView.geometry(r).segments[0];return [(seg[0][0]+seg[1][0])/2,(seg[0][1]+seg[1][1])/2]}''', record)
    # Fit the source and selected pipe to keep the test independent of auto-fit.
    page.evaluate('''()=>{const ps=ModuleHAttributes.getState().annotations.map(a=>ModuleHView.projectPoint([a.x,a.y]));
      ps.push(...ModuleHAttributes.getState().records.filter(r=>r.kind==='pipe').flatMap(r=>ModuleHView.geometry(r).segments.flat()));
      const xs=ps.map(p=>p[0]),ys=ps.map(p=>p[1]),h=document.querySelector('#cv').clientHeight;
      const scale=Math.min(750/(Math.max(...xs)-Math.min(...xs)),420/(Math.max(...ys)-Math.min(...ys)));
      __mf.view={scale,ox:Math.min(...xs)-150/scale,oy:Math.max(...ys)-(h-250)/scale};__hTest.draw();}''')
    capture_canvas(page, '#h-evidence-canvas')
    page.mouse.click(*screen(page, point))
    page.wait_for_function('label=>__mf.design.sel?.label===label', arg=record['label'])
    page.wait_for_function("__overlay.strokes.some(s=>s.dash.length===2 && s.color==='#ff3b45')")
    page.mouse.click(*screen(page, point), button='right')
    page.wait_for_selector('#ne-context:not(.hidden)')
    labels = page.locator('#ne-context button').all_text_contents()
    assert '배관속성 변경' in labels and '배관 분할' in labels and '참조 내경 변경' in labels
    page.locator('#ne-context button', has_text='참조 내경 변경').click()
    assert page.locator('#h-reference-picker').is_visible()
    start = len(h_ui[2])
    page.evaluate('''()=>{window.__referenceBusy=false;window.__referenceDesign=__mf.design;
      window.__busyObserver=new MutationObserver(()=>{if(!document.querySelector('#busy').classList.contains('hidden'))__referenceBusy=true;});
      __busyObserver.observe(document.querySelector('#busy'),{attributes:true,attributeFilter:['class']});}''')
    page.mouse.click(*screen(page, [2500, 100], project=True))
    page.wait_for_function('label=>__mf.design.tables.pipes.find(p=>p.label===label)?.dia===65', arg=record['label'])
    page.wait_for_function('label=>ModuleHAttributes.getState().records.find(r=>r.label===label)?.evidence.reference_annotation?.raw_text==="65A"', arg=record['label'])
    assert page.locator('#ha-iso').is_checked() == iso
    assert page.locator('#h-properties').is_hidden()
    page.wait_for_timeout(400)
    assert page.evaluate('__mf.design===__referenceDesign && !__referenceBusy')
    page.evaluate('__busyObserver.disconnect()')
    calls = h_ui[2][start:]
    assert calls.count(('POST', '/api/module-f/network-editor')) == 1
    assert not [p for method, p in calls if p in ('/api/module-f/design/preview', '/api/module-f/h-attributes/state') or method == 'GET' and p == '/api/module-f/network-editor']
    # A committed patch matches authoritative state, including the bulk-edit token.
    from routes.module_h_attributes import state as attributes
    authoritative = attributes(sess)
    assert page.evaluate('ModuleHAttributes.getState().revision') == authoritative['revision']
    # Native anchors are typed tuples in Python and JSON arrays on the wire.
    assert page.evaluate('ModuleHAttributes.getState().records') == json.loads(json.dumps(authoritative['records']))
    row = next(p for p in sess['design']['tables'].pipes if p['label'] == record['label'])
    assert row['inner_mm'] > 65
    page.keyboard.press('Control+z')
    page.wait_for_function('label=>__mf.design.tables.pipes.find(p=>p.label===label)?.dia===40', arg=record['label'])
    assert page.evaluate('__mf.stage') == 'design'
    assert page.locator('#ha-iso').is_checked() == iso
    page.keyboard.press('Control+Shift+z')
    page.wait_for_function('label=>__mf.design.tables.pipes.find(p=>p.label===label)?.dia===65', arg=record['label'])
    page.keyboard.press('Escape')
    page.wait_for_function('!__mf.design.sel')
    capture_canvas(page, '#h-attributes-canvas')
    page.evaluate("__mf.view.ox-=1;__hTest.draw()")
    page.wait_for_timeout(200)
    assert not page.evaluate("__overlay.strokes.some(s=>s.color==='#ff3b45')")
    page.wait_for_function('document.querySelectorAll("#hp-rows .hp-selected").length===0')
    page.screenshot(path=str(tmp / f'source-reference-{iso}.png'))


def test_evidence_all_and_selected_only_no_legacy_coloured_overlay(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    capture_canvas(page, '#h-evidence-canvas')
    page.check('#ha-show')
    page.wait_for_function('__overlay.strokes.some(s=>s.dash.length)')
    all_links = page.evaluate('__overlay.strokes.filter(s=>s.dash.length).length')
    assert all_links >= 2
    page.uncheck('#ha-show')
    page.wait_for_function('__overlay.strokes.length===0')
    page.evaluate("ModuleHProperties.select('pipe',ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.nominal_mm===40&&!r.generated).label)")
    page.wait_for_function('__overlay.strokes.some(s=>s.dash.length)')
    assert page.evaluate('__overlay.strokes.filter(s=>s.dash.length).length') < all_links
    assert not page.evaluate("__overlay.texts.some(t=>/^A\\d|^B\\d/.test(t[0]))")
    capture_canvas(page, '#h-attributes-canvas')
    page.check('#ha-show')
    page.wait_for_function("__overlay.strokes.some(s=>s.color==='#ff3b45')")
    assert not page.evaluate("__overlay.strokes.some(s=>['#5b9df1','#e0ae57','#b79af5'].includes(s.color))")


@pytest.mark.parametrize('iso', [False, True])
def test_reference_dashes_are_clearer_without_changing_evidence(h_ui, monkeypatch, iso):
    page, _, tmp = prepare(h_ui, monkeypatch)
    if iso:
        page.check('#ha-iso')
        settle(page)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    start = len(h_ui[2])
    capture_canvas(page, '#h-evidence-canvas')
    page.check('#ha-show')
    page.evaluate('__mf.view.ox-=1;__hTest.draw()')
    page.wait_for_function("__overlay.strokes.some(s=>s.color==='#b4cbd8' && s.path.length===2)")
    normal = page.evaluate("__overlay.strokes.filter(s=>s.color==='#b4cbd8' && s.path.length===2)")
    assert all(s['width'] == 1.5 and s['alpha'] == pytest.approx(.8)
               and s['dash'] == [6, 6] for s in normal)
    page.uncheck('#ha-show')
    page.evaluate("ModuleHProperties.select('pipe',ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.nominal_mm===40&&!r.generated).label)")
    page.wait_for_function("__overlay.strokes.some(s=>s.color==='#ff3b45' && s.path.length===2)")
    selected = page.evaluate("__overlay.strokes.filter(s=>s.color==='#ff3b45' && s.path.length===2)")
    assert all(s['width'] == pytest.approx(2.2) and s['alpha'] == 1 and s['dash'] == [6, 6] for s in selected), selected
    page.evaluate('__mf.view.scale*=1.1;__hTest.draw()')
    page.wait_for_timeout(200)
    assert page.evaluate("__overlay.strokes.some(s=>s.color==='#ff3b45' && Math.abs(s.width-2.2)<.001)")
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']
    page.screenshot(path=str(tmp / f'clearer-reference-dashes-{iso}.png'))


def test_existing_plan_split_picker_keeps_native_coordinates_and_cancel_is_read_only(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    point = screen(page, [5000, 2000])
    page.mouse.click(*point, button='right')
    page.locator('#ne-context button', has_text='배관 분할').click()
    label = page.evaluate('__mf.design.sel.label')
    before = len(sess['design']['tables'].pipes)
    page.click('#ne-pick-split')
    page.mouse.click(*point)
    assert page.locator('#ne-panel').is_visible()
    assert not page.locator('#ha-iso').is_checked()
    assert float(page.locator('#ne-length').input_value()) == pytest.approx(2)
    page.click('#ne-apply')
    page.wait_for_function('n=>__mf.design.tables.pipes.length===n', arg=before + 1)
    assert not page.locator('#ha-iso').is_checked()
    page.wait_for_function('n=>ModuleHAttributes.getState().records.filter(r=>r.kind==="pipe").length===n', arg=before + 1)
    assert page.evaluate('ModuleHView.planView().nodes.every(n=>Number.isFinite(n.x)&&Number.isFinite(n.y))')
    page.keyboard.press('Control+z')
    page.wait_for_function('n=>__mf.design.tables.pipes.length===n', arg=before)
    page.wait_for_function('n=>ModuleHAttributes.getState().records.filter(r=>r.kind==="pipe").length===n', arg=before)
    start = len(h_ui[2])
    page.evaluate('label=>ModuleHReferences.begin(label)', label)
    page.keyboard.press('Escape')
    assert page.locator('#h-reference-picker').is_hidden()
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']


def test_background_uses_original_opacity_not_full_colour_inference(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    result = page.evaluate('''async()=>{
      const s=__mf,c=document.querySelector('#cv').getContext('2d'),stroke=c.stroke.bind(c),calls=[];
      s.world.bundles=[{id:'blue',css:'#0000ff',segs:[0,0,1000,0]},{id:'yellow',css:'#ffff00',segs:[1000,0,1000,1000]}];
      c.stroke=(...a)=>{calls.push([c.strokeStyle,c.globalAlpha]);return stroke(...a);};
      __hTest.draw();await new Promise(requestAnimationFrame);await new Promise(requestAnimationFrame);c.stroke=stroke;return calls;
    }''')
    for color in ['#0000ff', '#ffff00']:
        alphas = [a for c, a in result if c == color]
        assert alphas and all(a == pytest.approx(.22) for a in alphas)


def test_compacted_pipe_keeps_each_original_segment_reference(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    result = page.evaluate('''()=>{
      const refs=[
        {annotation_id:'source-one',text_mm:40,text_xy_mm:[100,200],text_raw:'40A'},
        {annotation_id:'source-two',text_mm:40,text_xy_mm:[300,400],text_raw:'40'}];
      const record={evidence:{text_mm:40,evidence_chain:[refs[0],{evidence_chain:[refs[1],refs[0]]}]}};
      record.evidence.evidence_chain.push(record.evidence); // harmless shared/cyclic object guard
      return ModuleHReferences.sources(record).map(a=>[a.identity,a.x,a.y]).sort();
    }''')
    assert result == [['source-one',100,200],['source-two',300,400]]
