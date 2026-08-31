# -*- coding: utf-8 -*-
"""모듈 F 라우트 안전망 — 리팩터링 전에 먼저 깔았다.

**왜 이 파일이 필요했나.** 라우트 커버리지를 재 보니 모듈 F 의 64개 중
**15개만** 시험이 지나고 있었다(`scripts/_probe_f_route_coverage.py`). 그 상태로
`register()` 를 쪼개면, 옮긴 코드가 자유이름을 잃어 `NameError` 를 내도 아무도
모른다 — **등록만 보는 라우트 인벤토리 시험은 «실행할 때» 나는 오류를 못 잡는다.**
이 저장소가 이미 한 번 겪은 종류다.

두 겹으로 짠다. 겹마다 잡는 것이 다르고, **얕은 쪽이 깊은 쪽을 대신하지
못한다** — 그 한계를 여기 적어 둔다:

    ① 전수 두드리기 (얕다)   64개 전부를 세션 없이 한 번씩.
                              → 등록·데코레이터·import 단계의 깨짐을 잡는다.
                              → 대부분 `route_session` 에서 410 으로 막히므로
                                **핸들러 본문은 안 지난다.** 그것을 «덮였다» 고
                                읽으면 안 된다.
    ② 세션 하나로 깊게        진짜 도면으로 손질 상태까지 간 뒤, 세션만 있으면
                              도는 라우트들을 실제로 부른다.
                              → 본문의 `NameError`·`AttributeError` 를 잡는다.
"""
from __future__ import annotations

import importlib
import os
import sys
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
    return srv.app


def _client(app):
    c = app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def _rules(app) -> list:
    out = []
    for r in app.url_map.iter_rules():
        if "/api/module-f/" not in r.rule:
            continue
        for m in ("GET", "POST"):
            if m in r.methods:
                out.append((m, r.rule))
    return sorted(out)


def _concrete(rule: str) -> str:
    """경로 매개변수를 아무 값으로 — 없는 키에 대한 응답이 정상 동작이다."""
    if "<" not in rule:
        return rule
    return rule.split("<")[0] + "__none__"


# ═══════════════════════════════════════ ① 전수 두드리기 (얕다)
def test_모든_라우트가_예외를_안_던진다():
    """★이 시험이 «덮었다» 고 말하는 범위는 좁다.

    세션이 없으므로 대부분 `route_session` 이 410 으로 막고 **핸들러 본문은
    안 지난다.** 그래도 값이 있다: 데코레이터 조합이 깨졌거나, 모듈 수준
    import 가 죽었거나, 라우트가 등록만 되고 함수가 사라진 경우를 잡는다.

    깊은 쪽은 아래 세션 시험이 맡는다 — 둘 중 하나만으로는 부족하다.
    """
    app = _app()
    c = _client(app)
    rules = _rules(app)
    assert len(rules) >= 60, f"라우트가 갑자기 줄었다: {len(rules)}개"
    crashed = []
    for meth, rule in rules:
        path = _concrete(rule)
        try:
            rv = (c.get(path + "?sid=__none__") if meth == "GET"
                  else c.post(path, json={"sid": "__none__"}))
            if rv.status_code >= 500:
                crashed.append(f"{meth} {rule} → {rv.status_code}")
        except Exception as exc:  # noqa: BLE001 — 던지는 것 자체가 결함이다
            crashed.append(f"{meth} {rule} → {type(exc).__name__}: {exc}")
    assert not crashed, "라우트가 500 이거나 예외를 던진다:\n  " + \
                        "\n  ".join(crashed)


def test_세션이_없으면_통제된_실패로_답한다():
    """«없는 세션» 은 오류가 아니라 상태다 — 410 과 사람이 읽을 문장으로."""
    app = _app()
    c = _client(app)
    rv = c.get("/api/module-f/edit/state?sid=__none__")
    assert rv.status_code == 410
    body = rv.get_json() or {}
    assert body.get("ok") is False
    assert body.get("message"), "실패에 사람이 읽을 문장이 없다"


# ═══════════════════════════════════════ ② 세션 하나로 깊게
@pytest.fixture(scope="module")
def edited():
    """진짜 도면으로 손질 상태까지 — 한 번만 만들어 여러 시험이 나눠 쓴다.

    ★모듈 스코프다. 도면 한 장을 열고 채택·조립까지 가는 데 십여 초가 걸리므로
      시험마다 새로 만들면 전체 시간이 몇 배가 된다.
    """
    if not os.path.isfile(_DXF):
        pytest.skip(f"도면이 없다: {_DXF}")
    app = _app()
    c = _client(app)
    with open(_DXF, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, os.path.basename(_DXF)),
                         "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {}).get("sid")
    assert sid, f"열기 실패: {r.get_json()}"

    def wait(limit=3000):
        for _ in range(limit):
            j = c.get(f"/api/module-f/job?sid={sid}").get_json()
            if j.get("state") in ("done", "error", "idle"):
                return j
            time.sleep(0.1)
        return {"state": "timeout"}

    wait()
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
           .get("recon") or {})
    conf = (rec.get("adopt") or {}).get("conf_min") or 0.75
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True, "heads": {"conf_min": conf}})
    assert wait()["state"] == "done"
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    assert wait()["state"] == "done"
    return c, sid


# 세션만 있으면 도는 조회들. 여기 있는 것은 «본문까지» 지난다.
_READ_ONLY = [
    "/api/module-f/edit/state",
    "/api/module-f/world",
    "/api/module-f/recon",
    "/api/module-f/slot/state",
    "/api/module-f/sub/state",
    "/api/module-f/auto/state",
    "/api/module-f/merge/state",
    "/api/module-f/merge/modes",
    "/api/module-f/convert/result",
    "/api/module-f/design/preview",
    "/api/module-f/design/fitting-override",
    "/api/module-f/design/bore-override",
    "/api/module-f/job",
    "/api/module-f/saved",
    "/api/module-f/worst/reference-counts",
    "/api/module-f/auto/network-view",
    "/api/module-f/auto/handoff-hints",
]


@pytest.mark.parametrize("path", _READ_ONLY)
def test_조회_라우트가_세션_위에서_돈다(edited, path):
    """★얕은 두드리기가 못 보는 자리다 — 여기서부터 핸들러 본문이 돈다.

    실패는 «오류» 가 아니라 «상태» 일 수 있으므로(아직 표가 없다 등) 500 만
    막는다. 값이 무엇인지가 아니라 **코드가 끝까지 돌았는지** 를 본다.
    """
    c, sid = edited
    rv = c.get(f"{path}?sid={sid}")
    assert rv.status_code < 500, \
        f"{path} → {rv.status_code} · {rv.get_data(as_text=True)[:200]}"
    assert rv.get_json() is not None, f"{path} 가 JSON 이 아니다"


def test_손질_상태가_실제_망을_담고_있다(edited):
    """세션이 «있는 척» 만 하는 것이 아니라 진짜 망이 서 있어야 한다."""
    c, sid = edited
    st = (c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
          or {}).get("state") or {}
    assert st.get("counts", {}).get("pts", 0) > 100, f"절점이 너무 적다: {st}"
    assert st.get("counts", {}).get("edges", 0) > 100
    assert st.get("heads"), "헤드가 하나도 없다"


def test_자동_이음_훑기가_세션_위에서_돈다(edited):
    """[graph._autojoin_scan] 복잡도 46 짜리 함수 — 시험이 한 번도 안 지났다.

    붙일 후보가 없어도 «없다» 를 정상으로 답해야 한다.
    """
    c, sid = edited
    rv = c.post("/api/module-f/edit/autojoin/scan", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    body = rv.get_json() or {}
    assert "ok" in body


def test_최불리_선정과_되돌리기가_돈다(edited):
    """급수원 없이 부르면 «막힘» 이어야 한다 — 조용히 빈 값을 내면 안 된다."""
    c, sid = edited
    rv = c.post("/api/module-f/edit/worst", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    rv2 = c.post("/api/module-f/edit/worst-clear", json={"sid": sid})
    assert rv2.status_code < 500


def test_변환_입력칸_목록이_선다(edited):
    """세션과 무관하지만 엔진 부팅이 필요하다 — 부팅이 깨지면 여기서 걸린다."""
    c, _sid = edited
    body = c.get("/api/module-f/convert/fields").get_json() or {}
    assert body.get("ok") and body.get("groups"), body
    assert any(g.get("fields") for g in body["groups"])


# ═══════════════════════════════ ③ 표 확정까지 — 가장 큰 본문들을 지난다
#
# ★위의 «조회» 시험만으로는 부족하다. `design/preview` 는 표가 없으면 몇 줄
#   만에 «아직 확정 안 함» 으로 돌아 나온다 — 164줄짜리 본문은 안 지난다.
#   그 본문이야말로 좌표 변환·앵커 경로·표 직렬화가 모인 자리라 오류가
#   숨기 좋다. 그래서 진짜 표를 만든 세션을 하나 더 둔다.
@pytest.fixture(scope="module")
def confirmed(edited):
    """급수 시작 위치 원클릭 → 표 확정까지. 대명동으로 십여 초."""
    c, sid = edited

    def wait(limit=6000):
        for _ in range(limit):
            j = c.get(f"/api/module-f/job?sid={sid}").get_json()
            if j.get("state") in ("done", "error", "idle"):
                return j
            time.sleep(0.1)
        return {"state": "timeout"}

    st = (c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
          or {}).get("state") or {}
    heads = [(float(h[0]), float(h[1])) for h in (st.get("heads") or ())]
    hs = heads[::max(1, len(heads) // 40)][:40]
    groups = sorted((g.get("segs") or [] for g in st.get("body_groups") or []),
                    key=len, reverse=True)
    placed = False
    for s2 in groups[:12]:
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if len(pts) > 2000:
            pts = pts[::len(pts) // 2000]
        if not (pts and hs):
            continue
        best, bd = None, None
        for hx, hy in hs:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if bd is None or d < bd:
                best, bd = p, d
        if best is None or bd > 2000.0:
            continue
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": best[0], "y": best[1]})
        wait()
        w = ((c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
              or {}).get("state") or {}).get("worst") or {}
        if w.get("k"):
            placed = True
            break
    if not placed:
        pytest.skip("급수 시작 위치를 못 잡았다 — 도면 사정이지 코드 결함이 아니다")
    r = c.post("/api/module-f/design/build", json={"sid": sid})
    assert (r.get_json() or {}).get("ok"), r.get_json()
    assert wait()["state"] == "done"
    return c, sid


def test_수리계산_미리보기_본문이_끝까지_돈다(confirmed):
    """[api_design.module_f_design_preview · 164줄] 가장 큰 조회 본문이다.

    표가 없을 때는 몇 줄 만에 돌아 나오므로, 표를 만든 뒤에야 이 본문이
    지난다 — 좌표 변환 · 앵커 경로 되짚기 · 표 직렬화가 전부 여기 있다.
    """
    c, sid = confirmed
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    assert d.get("ok"), d
    v, t = d.get("view") or {}, d.get("tables") or {}
    assert v.get("nodes") and v.get("pipes"), "미리보기에 망이 없다"
    assert t.get("meta") and t.get("pipes"), "표가 비었다"
    # F-11c 가 실은 board 노드쌍 — 관경 덮기의 키다.
    assert any(p.get("ref") for p in v["pipes"]), "역참조가 하나도 없다"
    # ★앵커는 **단정하지 않는다.** 엔진의 `_anchor_node()` 가 무조건 None 을
    #   돌려주는 스텁이라 meta 의 「앵커 노드」는 항상 '?' 다(BLOCKED §30).
    #   여기서 「앵커가 있어야 한다」고 단정하면 시험이 «있지도 않은 기능» 을
    #   요구하게 된다 — 처음에 그렇게 적었다가 실패로 알았다.
    assert isinstance(v.get("anchor_path"), list), "형식은 지켜야 한다"


def test_앵커_경로가_아직_안_선다는_사실을_못_박는다(confirmed):
    """[BLOCKED §30] 지금은 «못 그리는» 것이 정상이다 — 그것을 기록으로 남긴다.

    엔진의 `_anchor_node()` 가 스텁이라 미리보기의 앵커 경로 계산이 통째로
    안 돈다. 이 시험은 **고쳐지면 실패한다** — 그때 이 시험과 BLOCKED §30 을
    함께 지우면 된다. 「안 되는 것을 안 된다고 아는 상태」를 지키는 장치다.
    """
    c, sid = confirmed
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    meta = dict((k, v) for k, v in ((d.get("tables") or {}).get("meta") or []))
    assert meta.get("앵커 노드") == "?", \
        f"앵커가 생겼다 — BLOCKED §30 과 이 시험을 지울 때다: {meta.get('앵커 노드')}"
    assert (d.get("view") or {}).get("anchor") is None


def test_확정된_표에서_조회_라우트가_다시_돈다(confirmed):
    """표가 생기면 «아직 없음» 분기 대신 진짜 분기가 돈다 — 그쪽도 지난다."""
    c, sid = confirmed
    for path in ("/api/module-f/design/preview",
                 "/api/module-f/design/fitting-override",
                 "/api/module-f/design/bore-override",
                 "/api/module-f/convert/result",
                 "/api/module-f/worst/reference-counts"):
        rv = c.get(f"{path}?sid={sid}")
        assert rv.status_code < 500, \
            f"{path} → {rv.status_code} · {rv.get_data(as_text=True)[:200]}"


def test_산출물_저장_본문이_돈다(confirmed, tmp_path_factory):
    """[design/emit] 파일을 실제로 쓴다 — 자산이 없으면 «통제된 실패» 여야 한다.

    ★성공을 단정하지 않는다. 이 시험이 도는 기계에 PIPENET 자산(SLF 템플릿)이
      없을 수 있다. 그때 500 이 아니라 사유 있는 실패로 답하는지가 요점이다.
    """
    c, sid = confirmed
    rv = c.post("/api/module-f/design/emit", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    body = rv.get_json() or {}
    if body.get("ok"):
        assert body.get("sdf", {}).get("bytes", 0) > 0
    else:
        assert body.get("message"), "실패에 사유가 없다"
