# -*- coding: utf-8 -*-
"""모듈 F 업로드 길 — «올린 뒤 도면이 뜨기까지» 의 계약을 시험이 지킨다.

여기서 지키는 것은 둘이다.

■ ① 압축 전송이 원본 전송과 **같은 결과**를 낸다
    화면이 큰 DXF 를 gzip 으로 보내기 시작했다(실측 B1F 110.6 → 14.2 MB).
    서버 `_save_upload` 는 종전에도 ".gz" 를 풀었지만 **통째로 RAM 에 들고**
    풀었다 — 흘려 쓰도록 고쳤다. 고친 뒤에도 저장된 바이트가 원본과 한 바이트도
    다르지 않아야 한다. 이 헬퍼는 모듈 F 만 쓰는 것이 아니라 업로드가 있는
    **모든 모듈**이 쓴다.

■ ② 도면은 잡이 «끝나기 전에» 그릴 수 있다
    `_open_job` 은 찍기판을 세우자마자 `sess["world"]` 를 앉히고, 그 뒤에
    정찰을 덤으로 돌린다. 그런데 화면은 잡이 다 끝나야 그렸다 —
    실측(B1F 110.6MB · 처음 여는 도면)으로 42초 중 **33초**가 «이미 서버에
    있는 도면» 을 안 그린 채 흘렀다. `_job_view` 의 `world_ready` 가 그
    사실을 화면에 말해 주는 신호다. 이 시험은 그 신호가 **잡이 도는 중에**
    참이 되는지 본다 — 끝난 뒤에만 참이면 고친 것이 없는 것과 같다.
"""
from __future__ import annotations

import gzip
import importlib
import io
import os
import sys
import threading
import time

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DXF = os.path.join(_ROOT, "samples", "dxf", "대명동201동 단위세대_layer정리.dxf")


def _app():
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    return srv


def _client(app):
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def _idle(c, sid, limit=6000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


# ═════════════════════════════════ ① 압축 전송 = 원본 전송
@pytest.mark.skipif(not os.path.isfile(_DXF), reason="시험 도면 없음")
def test_gzip_업로드가_원본과_같은_바이트를_남긴다():
    srv = _app()
    c = _client(srv.app)
    raw = open(_DXF, "rb").read()
    name = os.path.basename(_DXF)

    # 원본 그대로
    rv = c.post("/api/module-f/slot/open", data={
        "kind": "plan", "dxf_file": (io.BytesIO(raw), name)},
        content_type="multipart/form-data")
    assert rv.status_code == 200, rv.get_data(as_text=True)[:300]
    plain = rv.get_json()
    _idle(c, plain["sid"])
    saved_plain = srv.UPLOAD_DIR / plain["filename"]
    assert saved_plain.read_bytes() == raw

    # gzip 으로 — 화면이 큰 도면에 쓰는 길
    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb") as gz:
        gz.write(raw)
    rv = c.post("/api/module-f/slot/open", data={
        "kind": "plan", "dxf_file": (io.BytesIO(buf.getvalue()), name + ".gz")},
        content_type="multipart/form-data")
    assert rv.status_code == 200, rv.get_data(as_text=True)[:300]
    gzd = rv.get_json()
    _idle(c, gzd["sid"])

    # ★이름에서 ".gz" 가 떨어져야 한다 — 안 떨어지면 도면 «키» 가 달라지고,
    #   그 키로 저장한 찍은스펙이 원본 도면과 짝이 안 맞는다.
    assert gzd["filename"] == plain["filename"]
    assert (srv.UPLOAD_DIR / gzd["filename"]).read_bytes() == raw


def test_대상_파일이_열려_있어도_덮어쓴다():
    """★Windows 회귀 — «흘려 쓰기» 로 바꾸며 들어온 것.

    `.part` 로 받아 `os.replace` 로 제자리에 놓는데, Windows 는 대상 파일을
    다른 프로그램이 열어 두면 이름 바꾸기를 거부한다([WinError 5]). CAD
    뷰어로 그 도면을 열어 뒀거나 앞선 세션이 아직 읽고 있으면 그렇다 —
    종전(제자리 write_bytes)에는 되던 일이므로 그대로 두면 회귀다.

    실측으로 브라우저 검증 중에 바로 이 실패를 만났다
    («도면을 저장하지 못했습니다: [WinError 5] 액세스가 거부되었습니다»).
    """
    srv = _app()
    c = _client(srv.app)
    name = "__시험_열린채덮기__.dxf"
    dst = srv.UPLOAD_DIR / name
    dst.write_bytes(b"OLD")
    payload = b"NEW-CONTENT-" + os.urandom(64)
    holder = open(dst, "rb")          # 다른 프로그램이 열어 둔 상태
    try:
        rv = c.post("/api/module-f/slot/open", data={
            "kind": "plan", "dxf_file": (io.BytesIO(payload), name)},
            content_type="multipart/form-data")
        assert rv.status_code == 200, rv.get_data(as_text=True)[:300]
        assert dst.read_bytes() == payload, "열려 있다고 덮어쓰기를 건너뛰었다"
        assert not list(srv.UPLOAD_DIR.glob("*.part")), "반쪽 파일이 남았다"
    finally:
        holder.close()
        try:
            os.remove(dst)
        except OSError:
            pass


def test_압축폭탄은_푼_크기로_막힌다(monkeypatch):
    """`MAX_CONTENT_LENGTH` 는 «올라온» 바이트만 잰다 — 압축이면 그 자를 지난다.

    0 을 채운 20MB 는 gzip 으로 20KB 가 된다. 문 앞의 200MB 자로는 못 막는다.
    상한을 작게 갈아 끼워, 푼 크기를 세는 자가 실제로 작동하는지 본다.
    """
    srv = _app()
    c = _client(srv.app)
    monkeypatch.setattr(srv, "MAX_UNGZIP_BYTES", 64 * 1024)

    buf = io.BytesIO()
    with gzip.GzipFile(fileobj=buf, mode="wb") as gz:
        gz.write(b"0" * (4 * 1024 * 1024))     # 4MB → 압축본 수 KB
    rv = c.post("/api/module-f/slot/open", data={
        "kind": "plan", "dxf_file": (io.BytesIO(buf.getvalue()), "폭탄.dxf.gz")},
        content_type="multipart/form-data")
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    body = rv.get_json() or {}
    assert body.get("ok") is False
    assert "너무 큽니다" in (body.get("message") or ""), body
    assert not (srv.UPLOAD_DIR / "폭탄.dxf").exists()
    assert not list(srv.UPLOAD_DIR.glob("*.part")), "반쪽 파일이 남았다"


def test_깨진_gzip_은_500_이_아니라_사람이_읽을_실패다():
    """반쪽짜리 압축본은 «서버 오류» 가 아니라 «올린 것이 잘못» 이다."""
    srv = _app()
    c = _client(srv.app)
    broken = b"\x1f\x8b" + os.urandom(2048)
    rv = c.post("/api/module-f/slot/open", data={
        "kind": "plan", "dxf_file": (io.BytesIO(broken), "깨진.dxf.gz")},
        content_type="multipart/form-data")
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    body = rv.get_json() or {}
    assert body.get("ok") is False
    assert body.get("message")
    # ★반쪽 파일을 남기지 않는다 — 남으면 다음 열기가 그 반쪽을 읽는다.
    assert not (srv.UPLOAD_DIR / "깨진.dxf").exists()
    leftovers = [p.name for p in srv.UPLOAD_DIR.glob("*.part")]
    assert not leftovers, f"반쪽 파일이 남았다: {leftovers}"


# ═════════════════════════════════ ② 잡이 끝나기 전에 도면이 준비된다
@pytest.mark.skipif(not os.path.isfile(_DXF), reason="시험 도면 없음")
def test_정찰이_도는_동안_도면은_이미_받아갈_수_있다(monkeypatch):
    """정찰(덤)을 붙잡아 둔 채 «지금 도면 주세요» 가 되는지 본다.

    실도면에서는 이 창이 33초짜리라 눈으로도 보이지만, 시험은 크기에 기대면
    안 된다 — 정찰을 붙잡아 그 창을 명시적으로 만든다.
    """
    srv = _app()
    c = _client(srv.app)
    api_open = importlib.import_module("routes.module_f.api_open")

    hold = threading.Event()
    entered = threading.Event()

    def _slow_recon(sess, dxf, payload, t_open):
        entered.set()
        hold.wait(30)
        sess["recon"] = {"bundles": {}, "heads": [], "bands": {},
                         "elapsed_ms": 0}

    monkeypatch.setattr(api_open, "_recon_into", _slow_recon)

    rv = c.post("/api/module-f/slot/open", data={
        "kind": "plan", "dxf_file": (open(_DXF, "rb"), os.path.basename(_DXF))},
        content_type="multipart/form-data")
    assert rv.status_code == 200
    sid = rv.get_json()["sid"]
    try:
        assert entered.wait(120), "정찰까지 못 갔다 — 앞 단계에서 막혔다"

        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        assert j.get("state") == "run", "정찰을 붙잡았는데 잡이 안 돈다"
        # ★이 한 줄이 이 시험의 전부다.
        assert j.get("world_ready") is True, \
            "잡이 도는 중인데 world_ready 가 서지 않았다 — 화면이 도면을 " \
            "그릴 수 있는 시점을 알 길이 없어 정찰이 끝날 때까지 기다린다"

        w = c.get(f"/api/module-f/world?sid={sid}").get_json() or {}
        assert w.get("ok") is True, "world_ready 라 해 놓고 도면을 안 준다"
        assert w["world"]["counts"]["segs"] > 0
    finally:
        hold.set()
        _idle(c, sid)


def test_world_ready_는_잡이_붙기_전에도_거짓말을_안_한다():
    """세션은 있는데 잡이 아직 없을 때도 «없다» 고 말해야 한다."""
    from routes.module_f.jobs import _job_view
    assert _job_view({}) == {
        "state": "idle", "phase": "", "elapsed": 0.0, "lines": [],
        "world_ready": False}
    assert _job_view({"world": {"x": 1}})["world_ready"] is True
