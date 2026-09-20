"""Real browser clicks for review-before-download and failure handling."""
from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from urllib.parse import quote

import pytest

from test_module_f_viewport_browser import ui  # noqa: F401 (shared isolated browser)


@pytest.mark.parametrize("button,route,what", [
    ("btn-download-edit", "download", "kfp"),
    ("btn-download-worst-edit", "download", "worst-kfp"),
    ("mg-dl-kfp-edit", "merge/download", "kfp"),
])
@pytest.mark.parametrize("filename", ["editing_1cm.kfp", "B1F_편집용_1cm.kfp"])
def test_edit_download_requires_review_and_uses_server_filename(ui, button, route, what, filename):
    calls, dialogs = [], []

    def handle(req):
        query = parse_qs(urlparse(req.request.url).query)
        calls.append(query)
        assert query["variant"] == ["solver-edit"]
        assert query["what"] == [what]
        if "inspect" in query:
            req.fulfill(json={"ok": True, "source_sha256": "abc", "report": {
                "ok": True, "changed_nodes": 36, "max_coordinate_change_mm": 4.9,
                "length_changes": [{}] * 22, "max_length_change_mm": 8.389,
                "total_length_change_mm": -19.923}})
        else:
            assert query["source_sha256"] == ["abc"]
            header = ("attachment; filename=" + filename if filename.isascii()
                      else "attachment; filename*=UTF-8''" + quote(filename))
            req.fulfill(body=json.dumps({"grid_step": .01}), content_type="application/json",
                        headers={"Content-Disposition": header})

    ui.route(f"**/api/module-f/{route}?**", handle)
    ui.on("dialog", lambda d: (dialogs.append(d.message), d.accept()))
    ui.evaluate("""id => {
        __viewTest.setStage(id.startsWith('mg-') ? 'merge' : 'conv');
        document.getElementById(id).disabled = false;
    }""", button)
    with ui.expect_download() as event:
        ui.click("#" + button)
    download = event.value
    assert download.suggested_filename == filename
    assert json.loads(Path(download.path()).read_text()) == {"grid_step": .01}
    assert len(calls) == 2
    assert len(dialogs) == 1 and "8.389mm" in dialogs[0]
    assert "수리계산을 다시" in dialogs[0]


@pytest.mark.parametrize("case", ["blocked", "cancel", "changed"])
def test_failed_or_cancelled_review_never_downloads(ui, case):
    downloads, calls = [], []
    ui.on("download", lambda d: downloads.append(d))

    def handle(req):
        query = parse_qs(urlparse(req.request.url).query)
        calls.append(query)
        if "inspect" in query:
            req.fulfill(json={"ok": True, "source_sha256": "abc", "report": {
                "ok": case != "blocked", "errors": ["노드 겹침"], "changed_nodes": 1,
                "max_coordinate_change_mm": 4, "length_changes": [],
                "max_length_change_mm": 0, "total_length_change_mm": 0}})
        else:
            req.fulfill(status=409, json={"ok": False, "message": "원본이 바뀌었습니다"})

    ui.route("**/api/module-f/download?**", handle)
    ui.on("dialog", lambda d: d.dismiss() if case == "cancel" else d.accept())
    ui.evaluate("""() => {
        __viewTest.setStage('conv');
        document.getElementById('btn-download-edit').disabled = false;
    }""")
    ui.click("#btn-download-edit")
    if case == "blocked":
        ui.wait_for_function("document.getElementById('status').textContent.includes('노드 겹침')")
    elif case == "changed":
        ui.wait_for_function("document.getElementById('status').textContent.includes('원본이 바뀌었습니다')")
    ui.wait_for_function("document.getElementById('busy').classList.contains('hidden')")
    assert len(calls) == (2 if case == "changed" else 1)
    assert not downloads
