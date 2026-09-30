"""H route contracts: opt-in layout, unchanged F document and shared control IDs."""
from collections import Counter
from pathlib import Path
import re

import pytest
from flask import Flask, render_template

from routes.module_h import compose_document, register

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture()
def app():
    app = Flask(__name__, template_folder=str(ROOT / "templates"))
    app.testing = True
    register(app)
    return app


def test_h_preserves_every_engine_control_and_script(app):
    with app.app_context():
        original = render_template("module_f.html")
    response = app.test_client().get("/module-h")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    original_ids = Counter(re.findall(r'\bid="([^"]+)"', original))
    ids = Counter(re.findall(r'\bid="([^"]+)"', html))
    assert all(ids[key] == n for key, n in original_ids.items())
    assert not [key for key, n in ids.items() if n > 1]
    assert "MODULE H" in html and 'class="module-h"' in html
    scripts = re.findall(r'<script src="([^"]+)"', original)
    for src in scripts:
        suffix = r'&h_observer=[^"&]+' if src.startswith(('/static/module_f.js?', '/static/module_f_editor.js?')) else ''
        assert len(re.findall(r'<script src="' + re.escape(src) + suffix + '"', html)) == 1
    assert html.index('module_h_references.js') < html.index('module_h_evidence.js')
    assert html.index("module_h.js") > html.index("module_f.js")
    assert html.index("module_h_activity.js") < html.index("module_f.js")
    assert 'id="h-activity"' in html
    assert 'id="h-primary"' in html
    assert 'id="h-workflow"' not in html
    assert "저장 도면은 F와 공유" in html
    assert "module_h" not in original
    assert app.test_client().get('/api/module-h/selection-policy').json == {
        'ok':True,'selection_mode':'area_all','version':1}


def test_layout_contract_rejects_missing_or_duplicate_anchor():
    base = (ROOT / "templates/module_f.html").read_text(encoding="utf-8")
    for changed in (base.replace("</head>", ""), base + "</body>"):
        with pytest.raises(RuntimeError, match="contract changed"):
            compose_document(changed, "")


def test_main_application_registers_h_and_auth_gate(monkeypatch):
    # Same application factory/import used by the existing route regression suite.
    from test_module_f_routes import _app
    app = _app()
    client = app.test_client()
    response = client.get("/module-h")
    assert response.status_code in (302, 401, 403)
    assert client.get("/api/module-h/progress?operation=abcdefgh12345678").status_code in (302, 401, 403)
    assert client.get('/api/module-h/access-state?sid=missing').status_code in (302,401,403)
    with client.session_transaction() as session:
        session["authed"] = True
    assert client.get("/module-h").status_code == 200
    preview = client.get("/api/module-h/progress?operation=abcdefgh12345678")
    assert preview.status_code == 200 and preview.json["events"] == []
    home = client.get("/")
    assert home.status_code == 200
    home_html = home.get_data(as_text=True)
    card = re.search(r'<a\b[^>]*href="/module-h"[^>]*>(.*?)</a>', home_html, re.S)
    assert card is not None
    assert "Sprinkler Design Workbench (Simplified)" in card.group(1)
    assert "A drawing-centered workspace" in card.group(1)
    assert "existing validation checks" in card.group(1)
    assert not re.search(r"[가-힣]", card.group(1))
    f = client.get("/module-f").get_data(as_text=True)
    assert "MODULE F" in f and "module_h.js" not in f
