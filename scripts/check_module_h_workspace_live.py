"""Read-only deployment check for the single-workspace H and untouched F."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import get_key
from playwright.sync_api import sync_playwright


def main() -> None:
    """Use normal authentication; never upload/open a saved drawing or mutate it."""
    root = Path(__file__).resolve().parents[1]
    password = os.environ.get("LOGIN_PASSWORD") or get_key(root / ".env", "LOGIN_PASSWORD")
    if not password:
        raise SystemExit("Configured login unavailable; no login attempt made.")
    base = "http://127.0.0.1:5051"
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        context = browser.new_context(viewport={"width": 1500, "height": 960})
        response = context.request.post(
            base + "/login", form={"password": password, "next": "/module-h"}, max_redirects=0,
        )
        if response.status not in (302, 303):
            raise SystemExit("Normal login not accepted; no retry made.")
        errors: list[str] = []
        writes: list[str] = []
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.on("request", lambda req: writes.append(req.url) if req.method not in ("GET", "HEAD") else None)
        response = page.goto(base + "/module-h")
        assert response and response.status == 200
        page.wait_for_function("document.body.classList.contains('h-ready')")
        assert page.locator("#h-layout-error").count() == 0
        policy = context.request.get(base + '/api/module-h/selection-policy')
        assert policy.status == 200 and policy.json()['selection_mode'] == 'area_all'
        assert page.locator('#ed-k').is_hidden() and page.locator('#ed-k').is_disabled()
        assert page.locator("#h-workflow").count() == 0
        assert page.locator("#steps").is_hidden()
        assert page.locator(".side").is_hidden()
        assert page.locator("#h-primary").inner_text() == "Access"
        assert page.locator("body").evaluate("el=>el.classList.contains('h-activity-ready')")
        assert page.locator("#h-activity").is_hidden()
        assert page.locator("#h-activity #busy").count() == 1
        assert page.locator(".h-conversion-mark").is_visible()
        assert page.locator(".h-conversion-mark .h-format").all_text_contents() == ["DXF", "SDF"]
        assert page.locator(".h-conversion-mark [data-unit]").count() == 5
        assert page.locator(".h-conversion-mark rect").count() == 4
        assert page.locator('.h-architecture-label').text_content() == 'Pipe Network'
        assert page.locator('.h-conversion-mark [data-unit=processor] text').count() == 0
        assert page.locator('.h-network-graph circle').count() == 9
        assert page.locator("#h-welcome h2,#h-welcome p").count() == 0
        assert page.locator("#h-progress-canvas").is_hidden()
        assert page.locator(".h-activity-head").get_attribute("data-h-drag-handle") == "true"
        assert page.locator('#hp-reference, #hp-source-canvas, #hp-source-details').count() == 0
        assert page.evaluate('!!window.ModuleHProperties && !!window.ModuleHWindows')
        assert page.evaluate('!!window.ModuleHView?.drawOriginal')
        assert page.evaluate('!!window.ModuleHView?.geometry && !!window.ModuleHView?.plainNetwork')
        assert page.evaluate('!!window.ModuleHReferences?.begin && !!window.ModuleHSourceSelection?.inspect')
        assert page.evaluate('!!ModuleHView.annotationLayout && !!moduleFNetworkEditor.applyReference && !!moduleFNetworkEditor.clearSelection')
        assert page.locator('#h-reference-picker').is_hidden()
        assert page.locator('#h-source-text-canvas').is_hidden()
        assert page.locator('#h-source-text-canvas').evaluate("n=>getComputedStyle(n).pointerEvents") == 'none'
        assert page.locator('#h-back').get_attribute('aria-keyshortcuts') is None
        assert page.locator('#ed-recalc-row').evaluate("n=>n.classList.contains('h-obsolete-selection')")
        assert page.evaluate("typeof moduleFNetworkEditor.resolveConflict==='function'")
        assert page.locator('#ha-history-conflict').is_hidden()
        assert '교차해도' in page.locator('#pk-crop-info').text_content()
        assert page.evaluate('''()=>{
            const z=ModuleFRegions.finish([[0,0],[100,100],[0,100],[100,0]],1,{enclose:true});
            return z.fill_rule==='enclosed' && ModuleFRegions.area(z)===5000
              && ModuleFRegions.contains(z,[50,20]) && ModuleFRegions.contains(z,[50,80])
              && !ModuleFRegions.contains(z,[10,50]);
        }''')
        assert page.locator('#ha-iso').count() == 1
        assert page.locator('#ha-original').locator('..').inner_text() == '원본 도면'
        assert page.locator('#ha-original').is_checked()
        assert page.locator('.ha-view-options label').all_text_contents() == ['등각투상', '원본 도면', '전체 테이블']
        assert not page.locator('#ha-table').is_checked()
        actions = page.locator('#h-actionbar').bounding_box()
        assert actions['x'] == 12 and actions['width'] < 500 and actions['height'] <= 50
        canvas = page.locator('#cv').bounding_box()
        assert abs(canvas['y'] + canvas['height'] - 960) < 1
        assert page.locator('.h-action-copy').is_hidden()
        assert page.locator('#he-fold').get_attribute('aria-expanded') == 'true'
        for window in ('h-drawer','h-properties','h-attributes-panel','h-evidence'):
            assert page.locator('#'+window+' .h-resize-edge').count() == 8
        assert page.evaluate("getComputedStyle(document.body).getPropertyValue('--accent').trim()") == "#5B8DEF"
        output = root / "data/module_h_workspace_20260928"
        output.mkdir(exist_ok=True)
        page.screenshot(path=str(output / "live-h.png"))
        # Page-local CSS probe only: no operation, upload, worker or drawing writes.
        assert page.evaluate("""()=>{
            const busy=document.querySelector('#busy');
            busy.classList.remove('hidden');
            try { return getComputedStyle(document.querySelector('#h-welcome')).display; }
            finally { busy.classList.add('hidden'); }
        }""") == "none"
        page.click("#h-primary")
        page.wait_for_function('!document.querySelector("#h-access-plan").disabled')
        assert '도면을 연결해 하나의 배관망으로' not in page.locator('#h-access').inner_text()
        assert page.locator('#h-access-merge').is_disabled()
        assert page.locator('#dxf').is_hidden()
        page.screenshot(path=str(output / 'live-h-access.png'))
        page.click('#h-access-plan')
        page.wait_for_function('!document.querySelector("#h-drawer").hidden')
        assert page.locator("#dxf").is_visible()
        assert '자동으로 업로드' in page.locator('#dxf').get_attribute('aria-description')
        assert page.locator('#btn-open').locator('..').is_hidden()
        assert page.locator('#h-primary').inner_text() != '불러오기'
        assert page.locator(".h-drawer-head").get_attribute("data-h-drag-handle") == "true"
        assert "도면 업로드" in page.locator("#h-drawer").inner_text()
        drawer=page.locator('#h-drawer')
        assert abs(drawer.bounding_box()['width']-360)<2
        assert abs(drawer.bounding_box()['height']-205)<2
        page.screenshot(path=str(output / 'live-h-upload-compact.png'))
        width=drawer.bounding_box()['width']
        edge=drawer.locator('[data-edge=e]').bounding_box()
        x,y=edge['x']+3,edge['y']+30
        page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+80,y,steps=4);page.mouse.up()
        assert drawer.bounding_box()['width'] > width+70
        page.screenshot(path=str(output / "live-h-open.png"))
        page.keyboard.press("Escape")
        page.click("#h-more")
        assert page.locator("#h-table").is_disabled()
        assert page.locator("#h-sizing").is_disabled()
        page.keyboard.press("Escape")
        response = page.goto(base + "/module-f")
        assert response and response.status == 200
        page.wait_for_function("document.body.classList.contains('mf-clean')")
        assert page.locator("#mf-layout-error").count() == 0
        assert page.locator("#steps > div").count() == 5
        assert page.locator("#dxf").is_visible()
        assert page.locator('script[src*="module_h"]').count() == 0
        assert page.evaluate('!window.ModuleHNative && !window.ModuleHAccess')
        assert not errors, errors
        assert not writes, "Viewing F/H must not change drawing data."
        print("Live H: compact 360x205 upload, upright annotation layout and reference patch API loaded; Access verified. F intact. No drawing writes or browser errors.")
        browser.close()


if __name__ == "__main__":
    main()
