"""Real pointer and keyboard gestures, canonical bulk writes, and one undo."""
from test_module_h_browser import h_ui
from test_module_h_bottom_properties_browser import prepare


def open_table(page):
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')


def test_shift_select_bulk_c_and_single_undo(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch);open_table(page)
    rows=page.locator('#hp-rows tr');labels=[rows.nth(i).get_attribute('data-label') for i in range(2)]
    rows.nth(0).locator('td').first.click()
    rows.nth(1).locator('td').first.click(modifiers=['Shift'])
    page.wait_for_function('ModuleHProperties.getSelections().length===2')
    assert page.locator('#hp-rows .hp-selected').count()==2
    page.select_option('#hp-field','c');page.fill('#hp-value','135');page.click('#hp-apply')
    page.wait_for_function('labels=>labels.every(l=>ModuleHProperties.getData()?.pipes.find(p=>p.label===l)?.c===135)',arg=labels)
    assert page.evaluate('ModuleHProperties.getSelections().length')==2
    page.click('#hp-undo')
    page.wait_for_function('labels=>labels.every(l=>ModuleHProperties.getData()?.pipes.find(p=>p.label===l)?.c===120)',arg=labels)
    page.screenshot(path=str(tmp/'multi-selection.png'))


def test_drag_range_then_inline_edits_all_selected(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch);open_table(page)
    rows=page.locator('#hp-rows tr');labels=[rows.nth(i).get_attribute('data-label') for i in range(3)]
    a=rows.nth(0).locator('td').first.bounding_box();b=rows.nth(2).locator('td').first.bounding_box()
    page.mouse.move(a['x']+15,a['y']+a['height']/2);page.mouse.down()
    page.mouse.move(b['x']+15,b['y']+b['height']/2,steps=12);page.mouse.up()
    page.wait_for_function('ModuleHProperties.getSelections().length===3')
    rows.nth(0).locator('[data-field=c]').fill('145');rows.nth(0).locator('[data-field=c]').press('Tab')
    page.wait_for_function('labels=>labels.every(l=>ModuleHProperties.getData()?.pipes.find(p=>p.label===l)?.c===145)',arg=labels)
    assert page.evaluate('ModuleHProperties.getSelections().length')==3


def test_shift_canvas_heads_and_table_selection_share_set(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch)
    heads=page.evaluate("ModuleHAttributes.getState().records.filter(r=>r.kind==='head')")
    for i,h in enumerate(heads):
        pos=page.evaluate('''xy=>{const b=document.querySelector('#cv').getBoundingClientRect();return [b.left+__mf.toScreenX(xy[0]),b.top+__mf.toScreenY(xy[1])]}''',h['xy'][0])
        if i:page.keyboard.down('Shift')
        page.mouse.click(*pos)
        if i:page.keyboard.up('Shift')
    page.wait_for_function('ModuleHProperties.getSelections().length===2')
    open_table(page);page.select_option('#hp-kind','node')
    assert page.locator('#hp-rows .hp-selected').count()==2
    page.select_option('#hp-field','flow');page.fill('#hp-value','90');page.click('#hp-apply')
    page.wait_for_function('ModuleHProperties.getData()?.nodes.filter(n=>n.head).every(n=>n.head_spec.flow_lmin===90)')
    assert [r['flow_lmin'] for r in sess['design']['tables'].nozzles]==[90,90]


def test_sub_pick_coordinates_hidden_but_preserved(h_ui):
    page,sess,requests,tmp=h_ui
    page.evaluate("__mf.sub.picks=[[833411.125,139847.75],[662056.5,175767.25]];document.querySelector('#sub-picks').textContent='833411,139847';document.dispatchEvent(new Event('change'))")
    page.wait_for_function("document.querySelector('#h-sub-pick-status')?.textContent.includes('지정됨')")
    assert page.locator('#sub-picks').is_hidden()
    assert '833411' not in page.locator('#h-sub-pick-status').text_content()
    assert page.evaluate('__mf.sub.picks')==[[833411.125,139847.75],[662056.5,175767.25]]


def test_fitting_table_multiselect_updates_existing_counts_and_undo(h_ui,monkeypatch):
    page,sess,tmp=prepare(h_ui,monkeypatch);open_table(page)
    page.select_option('#hp-kind','fitting')
    rows=page.locator('#hp-rows tr');assert rows.count()>=2
    labels=[rows.nth(i).get_attribute('data-label') for i in range(2)]
    before=page.evaluate('ModuleHProperties.getData().property_fittings')
    rows.nth(0).locator('td').first.click();rows.nth(1).locator('td').first.click(modifiers=['Shift'])
    page.select_option('#hp-field','count');page.fill('#hp-value','2');page.click('#hp-apply')
    page.wait_for_function('labels=>labels.every(l=>ModuleHProperties.getData()?.property_fittings.find(f=>f.label===l)?.count===2)',arg=labels)
    assert len(page.evaluate('ModuleHProperties.getData().property_fittings'))==len(before)
    page.click('#hp-undo')
    page.wait_for_function('rows=>JSON.stringify(ModuleHProperties.getData()?.property_fittings)===JSON.stringify(rows)',arg=before)
