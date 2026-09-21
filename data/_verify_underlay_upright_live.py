"""Read-only deployment check; does not open or change a user's drawing."""
import hashlib
import json
import re
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
base = 'http://192.168.0.105:5051'
html = urllib.request.urlopen(base + '/login', timeout=15).read().decode('utf-8')
hint = re.search(r'<span class="pw-value">([^<]+)<', html)
with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page()
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.goto(base + '/module-f')
    if page.query_selector('input[type=password]'):
        assert hint, 'Local login hint unavailable'
        page.fill('input[type=password]', hint.group(1))
        page.click('button[type=submit]')
        page.wait_for_load_state('load')
    page.wait_for_function('!!window.moduleFNetworkEditor')
    assert '/static/module_f.js?v=20260917-underlay-upright' in page.content()
    response = page.request.get(base + '/static/module_f.js')
    assert response.status == 200
    assert hashlib.sha256(response.body()).digest() == hashlib.sha256((root/'static/module_f.js').read_bytes()).digest()
    assert all(page.locator('#mg-under-' + kind).count() == 1 for kind in ('plan','system','machineroom'))
    assert not errors, errors
    result = dict(ok=True, url=page.url, assets_match=True, version='20260917-underlay-upright', page_errors=errors)
    (root/'data/direct_editor_validation/underlay_upright_local_server.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result))
    browser.close()
