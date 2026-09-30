"""Selecting network objects must not reopen the H property window."""

from test_module_h_browser import h_ui, menu, seed
from test_module_h_bottom_properties_browser import prepare
from test_module_h_selection_colour_browser import wait_red


def click_plan(page, x, y):
    point = page.evaluate('''([x,y])=>{
        const b=document.querySelector('#cv').getBoundingClientRect();
        return [b.left+__mf.toScreenX(x),b.top+__mf.toScreenY(y)];
    }''', [x, y])
    page.mouse.click(*point)


def assert_stays_hidden(page):
    # Cover multiple visibility-poll cycles, including delayed editor updates.
    page.wait_for_timeout(500)
    assert page.locator('#h-properties').is_hidden()


def test_plan_selection_stays_closed_but_highlights_and_explicit_open_sync(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    before = page.evaluate('JSON.stringify(__mf.design.tables)')
    start = len(h_ui[2])
    click_plan(page, 5000, 2000)
    wait_red(page, 5000, 2000)
    assert_stays_hidden(page)
    menu(page, '#h-table')
    page.wait_for_function("document.querySelector('#hp-rows .hp-selected')?.dataset.label===__mf.design.sel.label")
    click_plan(page, 10000, 2000)
    wait_red(page, 10000, 2000)
    page.wait_for_function("document.querySelector('#hp-rows .hp-selected')?.dataset.label===__mf.design.sel.label")
    assert page.locator('#h-properties').is_visible()
    page.click('#hp-close')
    for x in (5000, 10000):
        click_plan(page, x, 2000)
        wait_red(page, x, 2000)
        assert_stays_hidden(page)
    menu(page, '#h-table')
    page.wait_for_function("document.querySelector('#hp-rows .hp-selected')?.dataset.label===__mf.design.sel.label")
    page.keyboard.press('Escape')
    click_plan(page, 5000, 4000)
    assert page.evaluate('__mf.design.sel.kind') == 'node'
    assert_stays_hidden(page)
    menu(page, '#h-table')
    page.wait_for_function("document.querySelector('#hp-kind').value==='node' && document.querySelector('#hp-rows .hp-selected')?.dataset.label===__mf.design.sel.label")
    page.click('#hp-fold')
    click_plan(page, 10000, 2000)
    wait_red(page, 10000, 2000)
    assert page.locator('#h-properties').bounding_box()['height'] <= 50
    assert page.evaluate('JSON.stringify(__mf.design.tables)') == before
    assert not [r for r in h_ui[2][start:] if r[0] == 'POST']


def test_closed_integrated_table_does_not_reopen_on_new_selection(h_ui):
    from test_module_f_sizing import merged, tables
    page, sess, requests, _ = h_ui
    sess['merged'] = merged(tables())
    sess['merge_summary'] = {'merged': True}
    seed(page, sess, 'merge')
    page.evaluate('async()=>{await __hTest.mergeState();await __hTest.merged();await moduleFNetworkEditor.refresh();}')
    page.wait_for_selector('#hp-rows tr')
    start = len(requests)
    page.click('#hp-close')
    for kind, label in [('mgpipe', '1'), ('mgpipe', '2'), ('mgnode', 'a')]:
        page.evaluate('([kind,label])=>__hTest.select(kind,label)', [kind, label])
        assert_stays_hidden(page)
    menu(page, '#h-table')
    page.wait_for_selector('#hp-rows tr[data-label="a"].hp-selected')
    page.keyboard.press('Escape')
    page.evaluate("__hTest.select('mgpipe','1')")
    assert_stays_hidden(page)
    assert not [r for r in requests[start:] if r[0] == 'POST']
