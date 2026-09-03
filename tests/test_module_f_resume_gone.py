# -*- coding: utf-8 -*-
"""«이어서 열기» — 원본 DXF 와 저장본의 **수명**.

여기서 지키는 것은 둘이다.
  ■ 애초에 지우지 않는다 — 저장본이 아직 쓰는 원본은 청소부가 건너뛴다.
  ■ 그래도 없으면(딴 데 있던 원본이 옮겨진 경우 등) 문 앞에서 곱게 거절한다.

─────────────────────────────────────────────────────────────────────────
업로드 폴더는 24시간이 지나면 정리된다(`_sweep_old_upload_files`). 찍은스펙과
표시캐시는 그대로 남지만 **원본 없이는 못 연다** — 실측으로 표시캐시가 있는
키도 엔진이 `SystemExit: DXF를 못 찾음` 으로 죽는다.

종전에는 그 죽음이 잡 안에서 일어나, 사람은 진행바가 «실패» 로 바뀌고
`SystemExit: …` 이라는 문장만 보았다. 왜 못 여는지도, 어떻게 되살리는지도
없었다. 실측: 저장본 11개 중 8개가 이 상태였다.

이 시험이 지키는 것은 둘이다.
  ① 문 앞에서 막는다 — 잡을 띄우지 않고 그 자리에서 거절한다.
  ② 사유와 «되살리는 길» 을 문장에 담는다.
"""
from __future__ import annotations

import importlib
import json
import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _app():
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    return srv.app


def _client(app):
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


@pytest.fixture()
def spec(tmp_path):
    """«원본이 사라진» 저장본 하나를 만든다 — 끝나면 걷어낸다.

    실제 작업폴더에 쓰는 이유: `_saved_keys()` 가 거기를 읽는다. 이름 앞에
    시험 표시를 달고 finally 로 반드시 지운다 — 데스크톱 G 와 같은 폴더라
    남기면 그쪽 목록에 낀다.
    """
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff
    key = "__시험_원본없음__"
    path = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"source_dxf": str(tmp_path / "없는도면.dxf")}, f,
                  ensure_ascii=False)
    try:
        yield key
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def test_원본이_없으면_목록이_그_사실을_싣는다(spec):
    c = _client(_app())
    items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
    row = next((it for it in items if it["key"] == spec), None)
    assert row is not None, "만든 저장본이 목록에 없다"
    # ★목록에서 «빼지» 않는다 — 빠지면 「내가 찍어 둔 것이 사라졌다」 로 읽힌다.
    assert row["source_exists"] is False
    assert row["source_dxf"], "어디를 찾았는지가 없으면 되살릴 수가 없다"


def test_원본이_없으면_잡을_띄우지_않고_사유를_말한다(spec):
    c = _client(_app())
    rv = c.post("/api/module-f/reopen", json={"key": spec})
    assert rv.status_code == 409, rv.get_data(as_text=True)[:200]
    body = rv.get_json() or {}
    assert body.get("ok") is False
    msg = body.get("message") or ""
    # 왜 안 되는지 · 어떻게 되살리는지 — 둘 다 있어야 한다.
    #
    # ★사유를 «24시간 정리» 로 단정하지 않는다. 청소부가 저장본이 가리키는
    #   원본을 이제 보호하므로(`_referenced_upload_files`), 그 문장은 대개
    #   **틀린 원인**을 통보하는 것이 된다. 시험도 그 옛 문구를 강요하고
    #   있었다 — 지켜야 할 것은 「무엇이 없는지」와 「어떻게 되살리는지」다.
    assert "원본 도면 파일이 없어" in msg, msg
    assert "다시 올리" in msg, msg
    assert "24시간" not in msg, f"틀린 원인을 단정한다: {msg}"
    # 찾던 곳을 알려 줘야 사람이 파일을 되찾을 수 있다.
    assert "찾던 곳" in msg, msg
    # ★잡이 서면 안 된다. 서면 진행바가 돌다 «SystemExit» 으로 끝난다.
    assert "sid" not in body, "거절인데 세션을 만들었다"


def test_원본이_있는_저장본은_종전대로_받는다():
    """막는 자를 너무 넓게 잡지 않았나 — 성한 저장본은 그대로 열려야 한다."""
    c = _client(_app())
    items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
    ok = next((it for it in items if it.get("source_exists")), None)
    if ok is None:
        pytest.skip("원본이 남아 있는 저장본이 없다")
    rv = c.post("/api/module-f/reopen", json={"key": ok["key"]})
    assert rv.status_code == 200, rv.get_data(as_text=True)[:200]
    assert (rv.get_json() or {}).get("sid")


# ═══════════════════════════ 애초에 지우지 않는다
def test_저장본이_쓰는_원본은_청소부가_건너뛴다(tmp_path):
    """★사람의 노동이 사라지던 자리.

    찍은스펙은 사람이 배관망을 손으로 한 줄씩 찍어 만든 것이고, 다시 열 때
    형상은 **원본에서 되세운다**. 그런데 청소부는 나이만 보고 업로드칸을
    비웠다 — 저장했다고 믿은 것이 다음 날 안 열린다.

    실측(2026-09-03): 저장본 11건 중 7건이 못 여는 상태였고, 그중 4건은
    원본이 바로 이 업로드칸을 가리키다 지워진 것이었다.

    이 시험은 «나이는 넘겼지만 저장본이 아직 가리키는» 도면 하나와, 아무도
    안 쓰는 도면 하나를 나란히 두고 청소부를 부른다.
    """
    import time
    srv = importlib.import_module("대조 서버")
    from routes.module_f.common import _boot
    _boot()
    from routes.module_f import world

    up = srv.UPLOAD_DIR
    up.mkdir(parents=True, exist_ok=True)
    used = up / "__시험_저장본이쓴다__.dxf"
    idle = up / "__시험_아무도안쓴다__.dxf"
    used.write_bytes(b"0")
    idle.write_bytes(b"0")
    old = time.time() - (srv._DIR_TTL_SECONDS + 3600)   # 나이를 넘긴다
    os.utime(used, (old, old))
    os.utime(idle, (old, old))

    key = "__시험_원본보호__"
    spec = os.path.join(world.pick_store_dir(), f"{key}_찍은스펙.json")
    with open(spec, "w", encoding="utf-8") as f:
        json.dump({"source_dxf": str(used)}, f, ensure_ascii=False)
    try:
        srv._sweep_old_upload_files(up)
        assert used.exists(),             "저장본이 아직 쓰는 원본을 지웠다 — 사람이 찍어 둔 것이 죽는다"
        # 막는 자를 너무 넓게 잡지 않았는가: 아무도 안 쓰는 것은 지워야 한다.
        assert not idle.exists(), "청소를 아예 안 했다 — 업로드칸이 무한정 쌓인다"
    finally:
        for p in (used, idle):
            try:
                p.unlink()
            except OSError:
                pass
        try:
            os.remove(spec)
        except OSError:
            pass


def test_저장본_목록을_못_읽으면_아무것도_안_지운다(monkeypatch, tmp_path):
    """★어느 쪽으로 눕는가 — 못 지운 파일은 다음에 지우면 되고,
    잘못 지운 파일은 되돌릴 수 없다.

    저장본 칸을 못 읽는 상황(권한·경로 이상)에서 «보호대상 0개» 로 읽으면
    청소부는 **전부 지워도 된다** 고 판단한다. 그 오독이 가장 비싸다.
    """
    import time
    srv = importlib.import_module("대조 서버")
    from routes.module_f import world

    def _blow_up():
        raise OSError("저장본 칸을 못 읽는다")

    monkeypatch.setattr(world, "referenced_sources", _blow_up)

    up = srv.UPLOAD_DIR
    victim = up / "__시험_못읽을때__.dxf"
    victim.write_bytes(b"0")
    old = time.time() - (srv._DIR_TTL_SECONDS + 3600)
    os.utime(victim, (old, old))
    try:
        srv._sweep_old_upload_files(up)
        assert victim.exists(),             "저장본 목록을 못 읽었는데 지웠다 — 보호 여부를 «모르는» 상태다"
    finally:
        try:
            victim.unlink()
        except OSError:
            pass


def test_저장본_칸은_엔진_없이도_답한다(monkeypatch):
    """청소부는 도면 하나 물으려고 G 트리를 통째로 올릴 수 없다.

    종전 `_saved_keys` 는 `services.cad_import…handoff` 를 물었다 — 부팅 전
    맥락에서는 `services` 가 sys.path 에 없어 그대로 터진다. 그러면 방어가
    «정리를 통째로 건너뛰는» 쪽으로 누워 업로드칸이 무한정 쌓인다.

    소스에서 낱말을 찾지 않는다 — 주석에 적힌 «handoff» 를 코드로 오독한다.
    엔진을 실제로 막아 두고 답이 나오는지 본다.
    """
    from routes.module_f.common import _boot
    _boot()
    from routes.module_f import world
    # 이미 올라온 것도 걷어야 한다 — 안 그러면 캐시가 답해 버려 검사가 헛돈다.
    for name in [n for n in sys.modules if n == "services"
                 or n.startswith("services.")]:
        monkeypatch.delitem(sys.modules, name)
    monkeypatch.setitem(sys.modules, "services", None)   # import 하면 터진다
    with pytest.raises(ImportError):
        importlib.import_module("services.cad_import.pipeline.handoff")

    d = world.pick_store_dir()                 # ★엔진 없이 답해야 한다
    assert d.endswith("0단계_새찍기"), d
    assert isinstance(world.referenced_sources(), set)
