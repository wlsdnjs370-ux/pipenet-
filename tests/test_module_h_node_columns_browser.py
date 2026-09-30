"""The compact node grid omits XY without losing coordinates or Z editing."""
from copy import deepcopy

import pytest

from test_module_h_browser import h_ui, seed
from test_module_h_bottom_properties_browser import prepare
from test_module_f_sizing import merged, tables


@pytest.mark.parametrize('scope', ['design', 'merge'])
def test_node_grid_has_no_xy_and_height_edit_preserves_xy(h_ui, monkeypatch, scope):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    if scope == 'merge':
        combined = tables()
        for field in ('nodes', 'pipes', 'nozzles', 'fittings', 'equipment', 'meta'):
            setattr(combined, field, deepcopy(getattr(sess['design']['tables'], field)))
        sess['merged'] = merged(combined)
        sess['merge_summary'] = {'merged': True}
        seed(page, sess, 'merge')
        page.evaluate('async()=>{await __hTest.mergeState();await __hTest.merged();await moduleFNetworkEditor.refresh();}')
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')
    page.locator('#hp-kind').select_option('node')
    assert page.locator('#hp-head th').all_text_contents() == [
        '노드', 'Z (m)', '헤드', '방수량 (L/min)', 'K (L/min/√bar)', '최소 압력 (bar)', '높이 반영',
    ]
    assert page.locator('#hp-rows [data-field=x],#hp-rows [data-field=y]').count() == 0
    assert page.locator('#hp-rows tr').first.locator('td').count() == 7
    original = page.evaluate('ModuleHProperties.getData().nodes.find(r=>r.head)')
    assert original and len(original['xyz']) == 3
    label = original['label']
    row = page.locator(f'#hp-rows tr[data-label="{label}"]')
    before = page.evaluate('JSON.stringify(ModuleHProperties.getData().nodes.map(r=>[r.label,...r.xyz.slice(0,2)]))')
    z = original['xyz'][2] + .25
    row.locator('[data-field=z]').fill(str(z))
    with page.expect_request(lambda req: req.method == 'POST' and '/api/module-f/network-editor' in req.url) as sent:
        row.get_by_role('button', name='반영', exact=True).click()
    command = sent.value.post_data_json['command']
    assert command['op'] == 'move_node'
    assert [command['x'], command['y']] == original['xyz'][:2]
    page.wait_for_function('({label,z})=>ModuleHProperties.getData()?.nodes.find(r=>r.label===label)?.xyz[2]===z', arg={'label': label, 'z': z})
    assert page.evaluate('JSON.stringify(ModuleHProperties.getData().nodes.map(r=>[r.label,...r.xyz.slice(0,2)]))') == before
    assert row.locator('[data-field=k]').is_enabled()
    assert row.locator('[data-field=pressure]').is_enabled()
    page.screenshot(path=str(tmp / f'node-z-only-{scope}.png'))
    page.click('#hp-undo')
    page.wait_for_function('({label,z})=>ModuleHProperties.getData()?.nodes.find(r=>r.label===label)?.xyz[2]===z', arg={'label': label, 'z': original['xyz'][2]})
    assert page.evaluate('JSON.stringify(ModuleHProperties.getData().nodes.map(r=>[r.label,...r.xyz.slice(0,2)]))') == before
    page.locator('#hp-kind').select_option('pipe')
    assert page.locator('#hp-head th').all_text_contents() == ['배관', '시작', '끝', '재질', '호칭경 (mm)', '내경 (mm)', '길이 (m)', 'C', '근거']
