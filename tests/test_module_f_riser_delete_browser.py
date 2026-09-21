"""[오너 2026-09-21] 실제 브라우저 — 세로관 삭제 카드(예/아니오), 하단 경고, 실제 크기 노드 점."""
import os

import pytest

from test_module_f_network_editor_browser import editor_ui, canvas_point, click_element  # noqa: F401 (fixture)
from test_module_f_viewport_browser import settle
from test_module_f_riser_delete import riser
from routes.module_f import network_edit as ne

IDLE = "document.querySelector('#busy').classList.contains('hidden')"


def open_merge(page, sess):
    from routes.module_f.api_merge import rebuild_merged
    page.click('#ne-close')  # the fixture leaves an extension card open in the design view
    sess['slots']['system']['riser'] = riser()
    sess['supply_mode'] = 'lsp_gravity'
    rebuild_merged(sess)
    page.evaluate("async()=>{__viewTest.setStage('merge');document.querySelector('#mg-iso').checked=true;"
                  "await __editUi.merged();}")
    page.click('#btn-fit')
    settle(page)


def shot(page, name):
    if os.environ.get('RISER_DELETE_SCREENSHOTS'):
        page.screenshot(path=os.path.join(os.environ['RISER_DELETE_SCREENSHOTS'], name))


def test_Delete_키_카드에서_예를_누르면_붙고_급수원은_화면에서_제자리(editor_ui):
    page, sess = editor_ui
    open_merge(page, sess)
    source = canvas_point(page, 'node', '1')
    page.evaluate("label=>__editUi.select('mgpipe',label)", 'r2')
    page.wait_for_function("document.querySelector('#ne-op').value==='resize'")
    page.keyboard.press('Delete')
    page.locator('#ne-join-yes').wait_for(state='visible')
    assert page.locator('.ne-join-q').inner_text() == '제거 후 상단과 하단 배관을 연결할까요?'
    assert '계통도 세로관' in page.locator('#ne-selected').inner_text()
    assert page.locator('#ne-op').is_hidden() and page.locator('#ne-apply').is_hidden()
    shot(page, 'riser_card.png')
    page.click('#ne-join-yes')
    page.wait_for_function(f"__mf.mergeSel?.label==='n2' && {IDLE}")
    settle(page)
    net = ne.ensure(sess, 'merge')['current']
    assert 'n3' not in net.nodes and 'r2' not in net.pipes
    assert net.nodes['1'].xyz[2] == pytest.approx(11.4) and net.nodes['10'].xyz[2] == pytest.approx(3.8)
    after = canvas_point(page, 'node', '1')
    assert abs(after['x']-source['x']) < 0.5 and abs(after['y']-source['y']) < 0.5
    assert page.locator('#ne-panel').is_hidden() and page.locator('#mg-split').is_hidden()


def test_우클릭_배관삭제_아니오는_배관만_지우고_하단에_경고_되돌리면_사라진다(editor_ui):
    page, sess = editor_ui
    open_merge(page, sess)
    click_element(page, 'pipe', 'r2', button='right')
    page.get_by_role('menuitem', name='배관 삭제', exact=True).click()
    page.locator('#ne-join-no').wait_for(state='visible')
    page.click('#ne-join-no')
    page.wait_for_function(f"!__mf.mergeView.pipes.some(p=>p.label==='r2') && {IDLE}")
    settle(page)
    assert 'n3' in ne.ensure(sess, 'merge')['current'].nodes
    assert page.locator('#mg-split').is_visible()
    assert '연결되지 않아 아직 산출할 수 없습니다' in page.locator('#mg-split').inner_text()
    assert '연결되지 않아' in page.locator('#status').inner_text()
    assert page.locator('#status').get_attribute('class') == 'warn'
    shot(page, 'riser_split.png')
    page.click('#ne-undo')
    page.wait_for_function(f"__mf.mergeView.pipes.some(p=>p.label==='r2') && {IDLE}")
    settle(page)
    assert page.locator('#mg-split').is_hidden()


def test_노드_점은_실제_크기라_줌에_따라_커지고_작아진다(editor_ui):
    page, sess = editor_ui
    open_merge(page, sess)
    probe = page.evaluate('''()=>{
      __mf.mergeSel=null;
      const c=document.querySelector('#cv'),ctx=c.getContext('2d'),dpr=c.width/c.clientWidth;
      const n=__mf.mergeView.nodes.find(n=>n.label==='n2'),cam=__mf.view;
      const px=(x,y)=>Array.from(ctx.getImageData(Math.round(x*dpr),Math.round(y*dpr),1,1).data.slice(0,3));
      const at=scale=>{
        cam.scale=scale;cam.ox=n.x-c.clientWidth/2/scale;cam.oy=n.y-c.clientHeight/2/scale;
        __viewTest.paint();
        const cx=c.clientWidth/2,cy=c.clientHeight/2,r=250*scale;
        return {r,ring:px(cx+r,cy),half:px(cx+r/2,cy),double:px(cx+2*r,cy)};
      };
      return {small:at(0.02),big:at(0.04),tiny:at(0.001)};
    }''')
    red = lambda p: p[0] > 150 and p[1] < 90 and p[2] < 90
    black = lambda p: max(p) < 40
    small, big = probe['small'], probe['big']
    assert small['r'] == pytest.approx(5) and big['r'] == pytest.approx(10)
    assert red(small['ring']) and black(small['double'])       # 5 px 반지름 고리
    assert red(big['ring']) and black(big['half'])             # 줌 2배 → 고리도 2배(옛 고리 자리는 점 안쪽)
    assert probe['tiny']['r'] < 0.6                             # 멀리서 보면 점을 안 그린다
