"""Border resize is UI-only, bounded, and available on every H floating window."""
from test_module_h_browser import h_ui, seed


def test_all_eight_borders_resize_and_reset_without_requests(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    page.click('#h-history')
    panel=page.locator('#h-drawer')
    assert panel.locator('.h-resize-edge').count()==8
    before=len([r for r in requests if r[0]=='POST'])
    for edge in ('n','e','s','w','ne','nw','se','sw'):
        panel.evaluate("n=>{n.style.cssText='left:440px!important;top:200px!important;width:450px!important;height:400px!important;right:auto!important;bottom:auto!important'}")
        box=panel.bounding_box()
        handle=panel.locator('[data-edge='+edge+']').bounding_box()
        x,y=handle['x']+handle['width']/2,handle['y']+handle['height']/2
        dx=-35 if 'w' in edge else 35 if 'e' in edge else 0
        dy=-30 if 'n' in edge else 30 if 's' in edge else 0
        page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+dx,y+dy,steps=4);page.mouse.up()
        after=panel.bounding_box()
        assert abs(after['width']-box['width']-abs(dx))<2,(edge,box,after)
        assert abs(after['height']-box['height']-abs(dy))<2,(edge,box,after)
    page.locator('.h-drawer-head h2').dblclick()
    assert panel.get_attribute('data-h-dragged') is None
    assert len([r for r in requests if r[0]=='POST'])==before
    # Newly created property/representative/bulk tables use the same manager.
    for target in ('h-properties','h-evidence','h-attributes-panel'):
        assert page.locator('#'+target+' .h-resize-edge').count()==8


def test_resize_while_busy_and_small_viewport_stays_reachable(h_ui):
    page,sess,_,_=h_ui
    seed(page,sess)
    page.click('#h-history')
    page.evaluate("__hTest.busy(true,'resize fixture')")
    panel=page.locator('#h-drawer');handle=panel.locator('[data-edge=e]').bounding_box()
    before=panel.bounding_box()['width']
    x,y=handle['x']+3,handle['y']+30
    page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+130,y,steps=4);page.mouse.up()
    assert panel.bounding_box()['width']>before+100
    page.set_viewport_size({'width':390,'height':640})
    page.wait_for_timeout(100)
    box=panel.bounding_box()
    assert box['x']>=7 and box['x']+box['width']<=383
    assert box['y']>=7 and box['y']+box['height']<=633
    page.evaluate("__hTest.busy(false)")
