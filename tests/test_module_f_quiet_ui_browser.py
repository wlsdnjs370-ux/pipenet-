"""Initial panels stay concise; collapsed controls retain their values/actions."""
from pathlib import Path
import os

import pytest

from test_module_f_viewport_browser import install_page,settle,pw_api,EDIT,WORLD


@pytest.fixture
def page():
    root=Path(__file__).resolve().parents[1]
    js=(root/'static/module_f.js').read_text(encoding='utf-8')
    js=js.replace('  setStage("open");\n  loadSaved();',
                  '  window.__quietUi={startNote,renderEdit};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch()
        p=browser.new_page(viewport={'width':1440,'height':1080})
        p.set_default_timeout(5000)
        errors=install_page(p,js)
        p.evaluate('''({world,edit})=>Object.assign(__mf,{sid:'quiet-ui',slot:'plan',method:'manual',world,edit})''',
                   dict(world=WORLD,edit=EDIT))
        yield p
        assert not errors,errors
        browser.close()


def test_resume_and_routine_adoption_explanations_stay_hidden(page):
    assert page.locator('#panel-resume').is_hidden()
    for stage in ['pick','edit']:
        page.evaluate('''stage=>{__quietUi.startNote('자동 인식 결과를 찍어 뒀습니다 — 채택 3개 · 유령 2개');__viewTest.setStage(stage);}''',stage)
        assert page.locator('#panel-start').is_hidden()
    assert page.locator('#ed-anchor-note').is_hidden()
    assert page.get_by_text('먼저 「헤드 고르기」',exact=False).is_hidden()
    # Failure information remains actionable, separate from routine explanation.
    page.evaluate("__quietUi.startNote('자동 인식 실패 — 직접 찍어 주세요.',true)")
    assert page.locator('#panel-start').is_visible()


def test_adoption_results_are_inside_auto_recommendations(page):
    page.evaluate("__viewTest.setStage('pick');document.querySelector('#pk-adopt-box').classList.remove('hidden')")
    assert page.locator('#pk-adopt-box').is_hidden()
    assert page.locator('#pk-show-low').is_hidden()
    page.locator('[data-fold="pk-auto-body"]').click()
    assert page.locator('#pk-adopt-box').is_visible()
    page.check('#pk-show-low')
    page.locator('[data-fold="pk-auto-body"]').click()
    assert page.locator('#pk-show-low').is_hidden() and page.locator('#pk-show-low').is_checked()


def test_edit_status_is_keyboard_accessible_and_keeps_fold_choice(page):
    page.evaluate("__viewTest.setStage('edit');__quietUi.renderEdit()")
    assert page.locator('#ed-info').is_hidden()
    button=page.locator('[data-fold="ed-status-body"]')
    button.focus();page.keyboard.press('Enter')
    assert button.get_attribute('aria-expanded')=='true'
    assert page.locator('#ed-info').is_visible()
    page.evaluate("__viewTest.setStage('pick');__viewTest.setStage('edit')")
    assert page.locator('#ed-info').is_visible()
    button.click()
    assert page.locator('#ed-info').is_hidden()


def test_calculation_details_and_conversion_options_are_collapsed_not_removed(page,tmp_path):
    page.evaluate("__viewTest.setStage('design')")
    for selector in ['#dg-summary','#dg-diagnose','#dg-bore-color','#dg-mk-dry','#dg-mk-swap','#dg-iso-note']:
        assert page.locator(selector).is_hidden(),selector
    assert page.locator('#dg-build').is_visible()
    assert page.locator('#dg-to-conv').is_visible()
    assert not page.locator('#dg-bore-color').is_checked()  # Corridor style is the default; evidence colors are opt-in.
    if os.environ.get('MODULE_F_QUIET_UI_SCREENSHOT'):
        page.screenshot(path=os.environ['MODULE_F_QUIET_UI_SCREENSHOT'],full_page=True)
    page.locator('[data-fold="dg-summary-body"]').click()
    assert page.locator('#dg-summary').is_visible() and page.locator('#dg-diagnose').is_visible()
    page.locator('[data-fold="dg-evidence-body"]').click()
    assert page.locator('#dg-bore-color').is_visible() and page.locator('#dg-mk-swap').is_visible()
    page.evaluate("__viewTest.setStage('conv')")
    assert page.locator('#conv-summary').is_hidden() and page.locator('#btn-conv-fields').is_hidden()
    page.locator('[data-fold="conv-options-body"]').click()
    assert page.locator('#conv-summary').is_visible()
    page.click('#btn-conv-fields')
    assert page.locator('#conv-modal').is_visible()
