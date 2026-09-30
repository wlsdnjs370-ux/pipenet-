"""Access transitions, actual server readiness and keyboard/small-screen QA."""
import pytest

from test_module_h_browser import h_ui, seed
from test_module_h_access import prepared_session


def settle_hub(page):
    page.wait_for_function('!document.querySelector("#h-access-plan").disabled')


@pytest.mark.parametrize('size',[(1500,960),(390,844),(844,390)])
def test_access_open_dock_expand_and_keyboard_without_upload(h_ui,size):
    page,sess,requests,tmp=h_ui
    page.set_viewport_size(dict(width=size[0],height=size[1]))
    assert page.locator('#h-welcome-open').inner_text()=='Access'
    page.click('#h-welcome-open');settle_hub(page)
    assert '도면을 연결해 하나의 배관망으로' not in page.locator('#h-access').inner_text()
    assert page.locator('#h-access-merge').is_disabled()
    assert page.locator('#h-access-add').is_disabled()
    assert page.locator('#dxf').is_hidden()
    page.screenshot(path=str(tmp/f'access-expanded-{size[0]}.png'))
    for id in ['h-access','h-access-plan','h-access-system','h-access-machineroom','h-access-merge']:
        box=page.locator('#'+id).bounding_box()
        assert box['x']>=0 and box['x']+box['width']<=size[0]+1
    page.click('#h-access-machineroom');settle_hub(page)
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    assert page.evaluate('__mf.slot')=='machineroom'
    assert page.locator('#dxf').is_visible()
    assert page.locator('#h-access').get_attribute('data-mode')=='docked'
    assert not page.locator('.h-access-card.is-ready').count()
    box=page.locator('#h-access').bounding_box()
    assert box['x']+box['width']<=size[0] and box['y']+box['height']<=size[1]
    page.click('#h-drawer-close')
    page.screenshot(path=str(tmp/f'access-docked-{size[0]}.png'))
    page.click('#h-access-size');settle_hub(page)
    assert page.locator('#h-access').get_attribute('data-mode')=='expanded'
    page.keyboard.press('Escape');settle_hub(page)
    assert page.locator('#h-access').get_attribute('data-mode')=='docked'
    assert not [r for r in requests if r[0]=='POST']


def test_actual_completion_dirty_revoke_and_http_failure(h_ui):
    page,sess,_,tmp=h_ui
    prepared=prepared_session()
    from routes.module_f import jobs
    sess['slots']=prepared['slots'];jobs._SESSIONS.pop(prepared['id'],None)
    seed(page,sess,'edit')
    page.evaluate('ModuleHAccess.refresh()')
    page.wait_for_function('ModuleHAccess.canMerge')
    assert page.locator('.h-access-card.is-ready').count()==3
    assert page.locator('#h-access-merge').is_enabled()
    page.screenshot(path=str(tmp/'access-ready.png'))
    sess['design']['h_attributes_dirty']=True
    page.evaluate('ModuleHAccess.refresh()')
    assert page.locator('#h-access-merge').is_disabled()
    sess['design']['h_attributes_dirty']=False
    page.route('**/api/module-h/access-state?*',lambda r:r.fulfill(status=500,body='failed'))
    page.evaluate('ModuleHAccess.refresh()')
    assert page.locator('#h-access-merge').is_disabled()
    assert '확인할 수 없습니다' in page.locator('#h-access-note').inner_text()
    page.unroute('**/api/module-h/access-state?*')
    page.evaluate('ModuleHAccess.refresh()')
    assert page.locator('#h-access-merge').is_enabled()


def test_reduced_motion_and_repeated_access_clicks(h_ui):
    page,_,requests,_=h_ui
    page.emulate_media(reduced_motion='reduce')
    page.evaluate('ModuleHAccess.open();ModuleHAccess.open()')
    settle_hub(page)
    assert page.locator('#h-access').is_visible()
    assert page.locator('#h-access').evaluate('el=>el.getAnimations({subtree:true}).length')==0
    page.click('#h-access-plan');settle_hub(page)
    assert page.evaluate('__mf.slot')=='plan'
    assert not [r for r in requests if r[0]=='POST']


def test_unfold_and_stagger_use_real_animations_and_busy_locks_cards(h_ui):
    page,_,requests,_=h_ui
    page.click('#h-welcome-open')
    assert page.locator('#h-access').evaluate("el=>el.getAnimations().some(a=>a.effect.getKeyframes().some(k=>k.clipPath))")
    assert page.locator('#h-access-plan').is_disabled()
    settle_hub(page)
    assert page.locator('#h-primary').inner_text()=='Access'
    page.click('#h-access-plan');settle_hub(page)
    page.evaluate("__hTest.busy(true,'검증 작업')")
    page.wait_for_function('document.querySelector("#h-access-plan").disabled')
    assert page.locator('#h-access-system').is_disabled()
    assert page.locator('#h-access-merge').is_disabled()
    page.evaluate('__hTest.busy(false)');settle_hub(page)
    assert not [r for r in requests if r[0]=='POST']


def test_real_slot_switch_add_remove_never_overwrites_plan(h_ui):
    page,sess,requests,_=h_ui
    plan=sess['design']
    seed(page,sess,'edit')
    page.evaluate('ModuleHAccess.refresh()')
    page.click('#h-access-system');settle_hub(page)
    assert sess['active']=='system'
    assert sess['slots']['plan']['design'] is plan
    assert page.evaluate('__mf.slot')=='system'
    assert page.locator('#dxf').is_visible()
    before=requests.count(('POST','/api/module-f/slot/add-system'))
    page.click('#h-access-add');settle_hub(page)
    assert requests.count(('POST','/api/module-f/slot/add-system'))==before+1
    assert sess['active']=='system2'
    assert page.evaluate('__mf.slot')=='system2'
    assert page.locator('#h-access-merge').is_disabled()
    page.once('dialog',lambda d:d.accept())
    page.click('[data-remove=system2]');settle_hub(page)
    assert sess['active']=='system' and 'system2' not in sess['system_extra']
    assert sess['slots']['plan']['design'] is plan
