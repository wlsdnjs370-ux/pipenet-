"""Read-only deployment smoke check in a new private browser, without opening drawings."""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import get_key
from playwright.sync_api import sync_playwright


def main() -> None:
    """Authenticate normally once; inspect F/H assets and initial screens only."""
    root = Path(__file__).resolve().parents[1]
    password = os.environ.get("LOGIN_PASSWORD") or get_key(root / ".env", "LOGIN_PASSWORD")
    if not password:
        raise SystemExit("Configured login unavailable; no login attempt made.")
    base = "http://127.0.0.1:5051"
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        context = browser.new_context(viewport={"width":1500,"height":960})
        result = context.request.post(base + "/login", form={"password":password,"next":"/module-f"}, max_redirects=0)
        if result.status not in (302,303):
            raise SystemExit("Normal login was not accepted; no retry made.")
        errors: list[str] = []
        writes: list[str] = []
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.on("request", lambda request: writes.append(request.url) if request.method not in ("GET","HEAD") else None)
        response = page.goto(base + "/module-f")
        assert response and response.status == 200
        page.wait_for_function("document.body.classList.contains('mf-clean')")
        assert page.locator("#mf-layout-error").count() == 0
        assert page.locator("#steps > div").count() == 5
        assert page.locator("#dxf").is_visible()
        for asset in ("module_f_layout.js", "module_f_layout.css"):
            assert context.request.get(base + "/static/" + asset).status == 200
        output = root / "data/module_f_ui_20260928"
        output.mkdir(exist_ok=True)
        page.screenshot(path=str(output / "live_f.png"))
        response = page.goto(base + "/module-h")
        assert response and response.status == 200
        page.wait_for_function("document.body.classList.contains('h-ready')")
        assert not page.locator("body").evaluate("n=>n.classList.contains('mf-clean')")
        assert page.locator("#h-layout-error").count() == 0
        assert page.locator("#h-workflow button").count() == 4
        assert not errors, errors
        assert not writes, "Initial F/H views must not issue data-changing requests."
        print("Live F: 5 stages, layout assets and upload control OK. H: separate UI OK. No drawing writes or browser errors.")
        browser.close()


if __name__ == "__main__":
    main()
