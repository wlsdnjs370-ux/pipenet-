"""Upright text, padded source bounds and transactional lightweight feedback."""
import pytest

from test_module_h_browser import h_ui, settle
from test_module_h_bottom_properties_browser import prepare
from test_module_h_reference_browser import screen
from test_module_h_source_text_browser import capture_text


@pytest.mark.parametrize('iso', [False, True])
def test_upright_multiline_rounded_bounds_and_matching_picker(h_ui, monkeypatch, iso):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    if iso:
        page.check('#ha-iso');settle(page)
    capture_text(page)
    layouts = page.evaluate('''()=>{
      const c=document.querySelector('#h-source-text-canvas').getContext('2d');
      const rows=[0,90,180,270].map((angle,i)=>({identity:'test-'+i,x:2500+i*2400,y:100,
        nominal_mm:150,raw_text:'SP 150\\n150',height_mm:120,rotation_deg:angle}));
      window.dispatchEvent(new CustomEvent('module-h-evidence-state',{detail:{...ModuleHAttributes.getState(),annotations:rows}}));
      return rows.map(a=>ModuleHView.annotationLayout(a,c));
    }''')
    page.wait_for_function('__sourcePaint.length===8')
    assert all(r['angle'] == pytest.approx(0) for r in page.evaluate('__sourcePaint'))
    for l in layouts:
        x, y, width, height = l['rect']
        assert x <= l['p'][0] - 6
        assert y <= l['p'][1] - l['size'] * .8 - 6
        assert width > l['size'] * 2 and height > l['size'] * 2 + 12
    # Real picker shares these same layouts and draws padded rounded rectangles.
    page.evaluate('''()=>{
      const c=document.querySelector('#h-reference-picker').getContext('2d'),old=c.roundRect.bind(c);
      window.__rounds=[];c.roundRect=(...a)=>{__rounds.push({args:a,dash:c.getLineDash()});return old(...a);};
      ModuleHReferences.begin(ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.nominal_mm===40&&!r.generated).label);
    }''')
    assert page.evaluate('__rounds.length') == 3
    assert all(r['args'][-1] == 6 and r['dash'] == [3, 3] for r in page.evaluate('__rounds'))
    # Select a natively vertical label at the RIGHT edge of its horizontal text.
    point = page.evaluate('''()=>{const a=ModuleHAttributes.getState().annotations.find(a=>a.rotation_deg===90),
      c=document.querySelector('#h-reference-picker').getContext('2d'),l=ModuleHView.annotationLayout(a,c),b=document.querySelector('#cv').getBoundingClientRect();
      return [b.left+l.rect[0]+l.rect[2]-7,b.top+l.p[1]-2];}''')
    page.mouse.click(*point)
    page.wait_for_function('document.querySelector("#h-reference-picker").hidden')
    assert any(p.get('bore_provenance', {}).get('reference_annotation', {}).get('rotation_deg') == 90 for p in sess['design']['tables'].pipes)
    page.screenshot(path=str(tmp / f'upright-reference-{iso}.png'))


def test_failed_reference_write_does_not_change_canvas_values_or_history(h_ui, monkeypatch):
    page, sess, _ = prepare(h_ui, monkeypatch)
    label = page.evaluate("ModuleHAttributes.getState().records.find(r=>r.kind==='pipe'&&r.nominal_mm===40&&!r.generated).label")
    page.evaluate('label=>ModuleHProperties.select("pipe",label)', label)
    page.wait_for_function('moduleFNetworkEditor.getState()?.sid===__mf.sid')
    before = page.evaluate('JSON.stringify([__mf.design.tables,ModuleHAttributes.getState(),moduleFNetworkEditor.getState().revision])')
    from routes.module_f import network_edit as ne
    monkeypatch.setattr(ne, 'save_history', lambda *_args: (_ for _ in ()).throw(OSError('fixture save failed')))
    start = len(h_ui[2])
    page.evaluate('label=>ModuleHReferences.begin(label)', label)
    page.mouse.click(*screen(page, [2500, 100], project=True))
    page.wait_for_function("document.querySelector('#h-reference-prompt').textContent.includes('fixture save failed')")
    assert page.evaluate('JSON.stringify([__mf.design.tables,ModuleHAttributes.getState(),moduleFNetworkEditor.getState().revision])') == before
    assert h_ui[2][start:].count(('POST', '/api/module-f/network-editor')) == 1
    assert page.locator('#busy').is_hidden()
    page.keyboard.press('Escape')
    assert page.locator('#h-reference-picker').is_hidden()
