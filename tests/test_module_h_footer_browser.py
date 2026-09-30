"""Compact H actions float bottom-left without reserving a full-width strip."""
import pytest

from test_module_h_browser import h_ui, seed, settle


@pytest.mark.parametrize('width',[1500,650,390])
def test_single_footer_without_coordinates_or_covering_strip(h_ui,width):
    page,sess,requests,tmp=h_ui
    page.set_viewport_size({'width':width,'height':960})
    seed(page,sess,'edit')
    assert page.locator('#coord').is_hidden()
    assert page.locator('#stage > .bar').is_hidden()
    controls=['btn-fit','h-back','h-history','h-primary']
    boxes=[]
    for control in controls:
        button=page.locator('#'+control)
        assert button.is_visible()
        assert button.evaluate("e=>e.parentElement.id")=='h-actionbar'
        boxes.append(button.bounding_box())
    for left,right in zip(boxes,boxes[1:]):
        assert left['x']+left['width']<=right['x']
        assert abs(left['y']-right['y'])<1
    assert boxes[0]['x']>=0
    assert boxes[-1]['x']+boxes[-1]['width']<=width
    canvas=page.locator('#cv').bounding_box()
    footer=page.locator('#h-actionbar').bounding_box()
    assert abs(canvas['y']+canvas['height']-960)<1
    assert 8<=footer['x']<=16 and 8<=960-footer['y']-footer['height']<=16
    assert footer['height']<=50
    assert footer['width'] <= sum(b['width'] for b in boxes)+34
    if width>=650:
        assert footer['width']<width*.7
    assert page.locator('.h-action-copy').is_hidden()
    # The formerly blocked bottom-right area now belongs to the drawing canvas.
    assert page.evaluate('''width=>document.elementFromPoint(width-10,950)?.id==='cv' ''', width)
    page.mouse.move(width/2,300);settle(page)
    assert page.locator('#coord').is_hidden()
    page.screenshot(path=str(tmp/f'module-h-footer-{width}.png'))
    assert not [path for method,path in requests if method=='POST']


def test_relocated_fit_restores_camera_even_while_processing(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'edit')
    page.click('#btn-fit');settle(page)
    fitted=page.evaluate('JSON.stringify(__mf.view)')
    data=page.evaluate('JSON.stringify([__mf.edit,__mf.world])')
    for busy in [False,True]:
        page.evaluate("__mf.view.ox+=12345;__mf.view.scale*=2;__hTest.draw()")
        if busy:
            page.evaluate("__hTest.busy(true,'테스트 처리 중')");settle(page)
        page.click('#btn-fit');settle(page)
        assert page.evaluate('JSON.stringify(__mf.view)')==fitted
        assert page.evaluate('JSON.stringify([__mf.edit,__mf.world])')==data
    page.evaluate('__hTest.busy(false)');settle(page)
    assert not [path for method,path in requests if method=='POST']
