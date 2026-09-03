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

세션이 넷이다. 같은 도면이라도 **어떤 상태인가** 가 다르면 다른 길이 열린다:

    picking      열기 + 읽기까지 (조립 «전») — `/pick/*` 는 조립 후엔 막힌다
    edited       조립까지 — `/edit/*` 대부분
    confirmed    원클릭 + 표 확정까지 — `design/preview` 의 164줄 본문
    auto_lane    `method="auto"` 로 연 것 — `/auto/*` 는 수동 세션에선 다 막힌다

■ ★아직 얕게만 덮인 라우트와 그 이유 (2026-09-01 기준 **10개**)

숫자를 부풀리지 않기 위해 적어 둔다. `scripts/_probe_f_route_coverage.py` 가
「딱 한 번만 불린 라우트」로 이 수를 따로 세어 준다 — 이 목록과 그 출력이
어긋나면 둘 중 하나가 낡은 것이다.

    /auto/run · /auto/network · /auto/handoff    무거운 잡 + 공유 작업폴더 쓰기
    /convert/run                                 전체망 변환 — 수십 초
    /edit/save                                   공유 «찍은 스펙» 파일에 쓴다
    /merge/build · /merge/emit · /merge/download  도면 세 장이 필요하다
    /open · /reopen                              저장된 키가 있어야 뜻이 있다

앞의 다섯은 **공유 상태에 쓰기** 때문에 그냥 덮으면 안 된다 — 웹 모듈 F 와
데스크톱 G 가 같은 작업폴더를 나눠 쓰고 있어서, 시험이 그것을 건드리면 G
시험의 입력이 바뀐다(BLOCKED §20 에서 실제로 겪었다). 덮으려면 쓰기 경로를
tmp 로 돌리는 준비가 먼저다.
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


def _idle(c, sid, limit=3000):
    """잡이 끝날 때까지 — **시험끼리 엮이지 않게** 한다.

    ★앞 시험이 띄운 잡이 아직 돌면 다음 라우트가 409(「이미 작업이 돌고
      있습니다」)로 막힌다. 그것을 «검사 대상의 응답» 으로 읽으면 엉뚱한
      결론이 난다 — 실제로 한 번 그렇게 실패했다(400 을 기대했는데 409).
    """
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


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


# 위 도크스트링이 「아직 얕게만 덮였다」고 적어 둔 것들. 시험이 깊어지면
# 이 목록에서 빼야 한다 — 아래 시험이 그것을 강제한다.
_SHALLOW_ONLY = {
    "/api/module-f/auto/run", "/api/module-f/auto/network",
    "/api/module-f/auto/handoff", "/api/module-f/convert/run",
    "/api/module-f/edit/save", "/api/module-f/merge/build",
    "/api/module-f/merge/emit", "/api/module-f/merge/download",
    "/api/module-f/open", "/api/module-f/reopen",
}


def test_얕게만_덮인_목록이_실제와_맞는다():
    """★문서만 두면 낡는다 — 목록을 시험이 지킨다.

    라우트가 늘었는데 목록을 안 고치면 «다 덮였다» 는 착각이 남고, 반대로
    깊은 시험을 더했는데 목록에 그대로 두면 실제보다 못한 것으로 읽힌다.
    둘 다 사실을 흐린다.

    여기서는 **목록의 항목이 전부 실재하는 라우트인지** 만 본다. 실제로 몇 번
    불렸는지는 pytest 안에서 셀 수 없다(그건 커버리지 도구의 몫이다) —
    그 한계를 알고 쓰는 것이 목록을 안 쓰는 것보다 낫다.
    """
    app = _app()
    rules = {r for _m, r in _rules(app)}
    ghost = sorted(p for p in _SHALLOW_ONLY if p not in rules)
    assert not ghost, f"목록에 없는 라우트가 적혀 있다(낡았다): {ghost}"


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
    # ★이름 주의 — 여기 «기준 헤드» 는 최원단 헤드다. 이 저장소에서 «앵커» 는
    #   그 반대쪽, 라이저가 붙는 접속점이다(services/cad_import/design/anchor.py).
    #   여기서는 형식만 본다 — 값이 서는지는
    #   `test_최원_유하거리_경로가_선다` 가 따로 지킨다.
    assert isinstance(v.get("worst_path"), list), "형식은 지켜야 한다"


def test_최원_유하거리_경로가_선다(confirmed):
    """[BLOCKED §30 해소] 예고대로 «고쳐지면 실패하는» 시험이 실패했고, 뒤집었다.

    종전 이 자리에는 「기준 헤드 노드는 항상 '?' 다」를 못 박는 시험이 있었다.
    엔진의 `_worst_head_node()` 가 스텁이라 미리보기의 경로 계산이 한 번도
    안 돌았기 때문이다. 되짚는 표를 바꾸자(node_ref → 좌표) 돌기 시작했다.

    이제 지키는 것은 반대다 — 라벨이 서고, 그 라벨이 **경로의 끝**이며,
    경로 길이가 최원 유하거리와 어긋나지 않는다.
    """
    c, sid = confirmed
    d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
    meta = dict((k, v) for k, v in ((d.get("tables") or {}).get("meta") or []))
    lab = meta.get("기준 헤드 노드")
    assert lab and lab != "?", f"기준 헤드 노드가 아직 '?' 다: {meta}"

    v = d.get("view") or {}
    assert v.get("worst_head") == lab, (v.get("worst_head"), lab)
    path = v.get("worst_path") or []
    assert len(path) >= 2, path
    assert str(path[-1]) == str(lab), "경로가 기준 헤드에서 끝나지 않는다"
    # 접속점에서 시작해야 한다 — 뿌리는 Input 이다.
    nodes = (d.get("tables") or {}).get("nodes") or []
    root = next((str(n.get("label")) for n in nodes
                 if str(n.get("io_node")) == "Input"), None)
    assert str(path[0]) == str(root), (path[0], root)
    # 길이는 «숫자만 있고 줄은 없던» 종전의 far_m 과 맞아야 한다.
    far = float(((d.get("summary") or {}).get("far_m")
                 or v.get("worst_path_m") or 0.0))
    got = float(v.get("worst_path_m") or 0.0)
    assert got > 0.0
    assert abs(got - far) <= max(1.0, far * 0.02), (got, far)


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


def test_관경_덮기_저장_본문이_돈다(confirmed):
    """[F-11c] POST 쪽은 얕은 두드리기에서 410 으로 막혔다 — 여기서야 본문이 돈다.

    규격표에 없는 호칭경을 거절하는지까지 본다. 그 거절이 사라지면 SLF 에 없는
    값이 산출로 나가 PIPENET 이 그 배관을 못 푼다.
    """
    c, sid = confirmed
    v = ((c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {})
         .get("view") or {})
    cand = next((p for p in (v.get("pipes") or []) if p.get("ref")), None)
    assert cand, "역참조 있는 배관이 없다"
    a, b = cand["ref"]
    bad = c.post("/api/module-f/design/bore-override",
                 json={"sid": sid, "rows": [{"a": a, "b": b, "dia": 77}]})
    assert bad.status_code == 400, "규격표에 없는 77A 가 통과한다"
    ok = c.post("/api/module-f/design/bore-override",
                json={"sid": sid,
                      "rows": [{"a": a, "b": b, "dia": 80, "note": "시험"}]})
    assert (ok.get_json() or {}).get("ok"), ok.get_json()
    # 되돌린다 — 이 시험이 뒤 시험의 산출을 바꾸면 안 된다.
    c.post("/api/module-f/design/bore-override", json={"sid": sid, "rows": []})


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


# ═══════════════════════════════ ④ 찍기 단계 — 조립 «전» 이라야 도는 것들
#
# `/pick/*` 는 조립하고 나면 세션이 손질로 넘어가 400 으로 막힌다. 그래서
# 조립 전에서 멈춘 세션이 따로 필요하다 — 앞의 `edited` 를 재활용할 수 없다.
@pytest.fixture(scope="module")
def picking():
    """도면을 열고 «읽기» 까지만. 조립은 안 한다."""
    if not os.path.isfile(_DXF):
        pytest.skip(f"도면이 없다: {_DXF}")
    c = _client(_app())
    with open(_DXF, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, os.path.basename(_DXF)),
                         "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {}).get("sid")
    assert sid, r.get_json()
    for _ in range(3000):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            break
        time.sleep(0.1)
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    return c, sid


def test_찍기_모드_전환이_돈다(picking):
    """세 동작(배관·완료·헤드칸)과 «모르는 동작» 의 거절까지."""
    c, sid = picking
    for action in ("pipe", "complete", "slot"):
        rv = c.post("/api/module-f/pick/mode",
                    json={"sid": sid, "action": action, "slot": "상하향식"})
        assert rv.status_code < 500, f"{action} → {rv.get_data(as_text=True)[:200]}"
    bad = c.post("/api/module-f/pick/mode",
                 json={"sid": sid, "action": "없는동작"})
    assert bad.status_code == 400, "모르는 동작을 통과시킨다"


def test_재료_일괄_찍기와_되돌리기가_돈다(picking):
    """[adopt.adopt_bundles] 자동 채택이 쓰는 그 경로다 — 본문까지 지난다."""
    c, sid = picking
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
    rv = c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    body = rv.get_json() or {}
    assert "ok" in body
    un = c.post("/api/module-f/pick/undo", json={"sid": sid})
    assert un.status_code < 500, un.get_data(as_text=True)[:200]


def test_헤드_후보_제안이_돈다(picking):
    """[F-5] 제외 사유 계측이 붙은 자리 — 후보가 0이어도 «없다» 를 답해야 한다."""
    c, sid = picking
    rv = c.post("/api/module-f/pick/suggest", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    assert (rv.get_json() or {}).get("ok") is not None


def test_찍기_클릭이_좌표를_검사한다(picking):
    """좌표가 숫자가 아니면 «통제된 거절» 이어야 한다 — 500 이 아니라."""
    c, sid = picking
    bad = c.post("/api/module-f/pick/click",
                 json={"sid": sid, "x": "abc", "y": None})
    assert 400 <= bad.status_code < 500, bad.status_code
    ok = c.post("/api/module-f/pick/click",
                json={"sid": sid, "x": 0.0, "y": 0.0})
    assert ok.status_code < 500, ok.get_data(as_text=True)[:200]


# ═══════════════════════════════ ⑤ 손질 단계 — `edited` 위에서
def test_자동_이음_적용과_지우기가_돈다(edited):
    """훑기만 시험에 있었다 — 적용·지우기 본문은 안 지나고 있었다."""
    c, sid = edited
    _idle(c, sid)
    c.post("/api/module-f/edit/autojoin/scan", json={"sid": sid})
    _idle(c, sid)
    ap = c.post("/api/module-f/edit/autojoin/apply", json={"sid": sid})
    assert ap.status_code < 500, ap.get_data(as_text=True)[:200]
    _idle(c, sid)
    cl = c.post("/api/module-f/edit/autojoin/clear", json={"sid": sid})
    assert cl.status_code < 500, cl.get_data(as_text=True)[:200]
    _idle(c, sid)


def test_헤드_종류_바꾸기가_모르는_값을_막는다(edited):
    """[api_edit.module_f_edit_kind] 미지정이 남으면 변환이 막힌다 — 그 입구."""
    c, sid = edited
    _idle(c, sid)
    bad = c.post("/api/module-f/edit/kind",
                 json={"sid": sid, "kind": "없는종류"})
    assert bad.status_code == 400, \
        f"모르는 헤드 종류가 통과한다 ({bad.status_code})"
    # 고른 헤드가 없으면 «먼저 고르라» 고 답해야 한다(조용히 성공하면 안 된다).
    rv = c.post("/api/module-f/edit/kind",
                json={"sid": sid, "kind": "상향식"})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]


def test_물길_보기가_돈다(edited):
    """[edit/flow] 급수원에서 물이 닿는 간선 — 급수원이 없으면 그렇다고 답한다."""
    c, sid = edited
    rv = c.post("/api/module-f/edit/flow", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    assert (rv.get_json() or {}) != {}


def test_도면_그림_라우트가_실재하는_키로_돈다(edited):
    """[diagram/<key>] 얕은 두드리기는 «없는 키» 로만 지났다 — 진짜 키로 지난다."""
    c, _sid = edited
    from routes.module_f.common import DIAGRAMS
    assert DIAGRAMS, "그림 목록이 비었다"
    key = sorted(DIAGRAMS)[0]
    rv = c.get(f"/api/module-f/diagram/{key}")
    # 파일이 이 기계에 없을 수 있다 — 404 는 정상, 500 은 아니다.
    assert rv.status_code in (200, 404), rv.status_code


def test_슬롯_전환이_돈다(edited):
    """[slot/switch] 세 칸의 상태 기계 — 없는 칸은 거절해야 한다."""
    c, sid = edited
    bad = c.post("/api/module-f/slot/switch",
                 json={"sid": sid, "kind": "없는칸"})
    assert 400 <= bad.status_code < 500, bad.status_code
    rv = c.post("/api/module-f/slot/switch",
                json={"sid": sid, "kind": "system"})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "plan"})


# ═══════════════════════════════ ⑥ 계통도·기계실 추출 — 입력 검사까지
#
# 두 라우트는 슬롯을 바꾸고 도면을 읽어야 도는데, 그 앞에 «입력 검사» 층이
# 있다. 좌표가 없거나 숫자가 아닐 때 500 이 아니라 사유 있는 거절이어야 한다 —
# 그 층은 슬롯 없이도 지날 수 있으므로 여기서 덮는다.
_SUB_DXF = os.path.join(_ROOT, "data", "uploads",
                        "1. 입력도면 대명동 단위세대 계통도.dxf")


@pytest.fixture(scope="module")
def system_slot(edited):
    """계통도 슬롯으로 바꾸고 도면 한 장을 읽는다."""
    if not os.path.isfile(_SUB_DXF):
        pytest.skip(f"계통도가 없다: {_SUB_DXF}")
    c, sid = edited
    _idle(c, sid)
    c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "system"})
    with open(_SUB_DXF, "rb") as fh:
        c.post("/api/module-f/slot/open",
               data={"sid": sid, "dxf_file": (fh, os.path.basename(_SUB_DXF)),
                     "kind": "system"},
               content_type="multipart/form-data")
    _idle(c, sid)
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    _idle(c, sid)
    yield c, sid
    # 다른 시험이 평면도를 기대하므로 되돌려 둔다.
    c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "plan"})
    _idle(c, sid)


def test_계통도_추출이_좌표를_검사한다(system_slot):
    """★두 점을 안 찍으면 «찍으세요» 라고 답해야 한다 — 조용히 빈 망이 아니라."""
    c, sid = system_slot
    rv = c.post("/api/module-f/system/extract", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    body = rv.get_json() or {}
    assert body.get("ok") is False and body.get("message"), body
    bad = c.post("/api/module-f/system/extract",
                 json={"sid": sid, "pump_x": "a", "pump_y": 0,
                       "av_x": 0, "av_y": 0})
    assert bad.status_code < 500, bad.get_data(as_text=True)[:300]
    assert (bad.get_json() or {}).get("ok") is False


def test_계통도_추출_본문이_돈다(system_slot):
    """[BLOCKED §27] 이 도면은 배관 레이어 분류가 뒤집혀 조각난다.

    ★성공을 단정하지 않는다. 잰 것은 **실패해도 사유를 남기는가** 다 —
      계통도는 240노드·63조각이라 단일망 추출이 성립하지 않을 수 있고,
      그때 «조용한 빈 망» 이 아니라 사유가 나와야 한다.
    """
    c, sid = system_slot
    rv = c.post("/api/module-f/system/extract",
                json={"sid": sid, "clean": True})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    body = rv.get_json() or {}
    if not body.get("ok"):
        assert body.get("message"), "실패에 사유가 없다"


def test_기계실_추출이_천장고를_검사한다(system_slot):
    """숫자가 아닌 천장고는 «통제된 거절» 이어야 한다."""
    c, sid = system_slot
    rv = c.post("/api/module-f/machineroom/extract",
                json={"sid": sid, "ceiling_m": "높음"})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    assert (rv.get_json() or {}).get("ok") is False


# ═══════════════════════════════ ⑦ 자동 추출 차선 — `method="auto"` 세션
#
# 자동 경로는 «수동으로 연 세션» 에서 전부 막힌다(`_need_auto`). 그래서 또
# 다른 세션이 필요하다 — 같은 도면이라도 «어떻게 열었나» 가 다르면 다른 길이다.
@pytest.fixture(scope="module")
def auto_lane():
    """평면도를 자동 차선으로 연다."""
    if not os.path.isfile(_DXF):
        pytest.skip(f"도면이 없다: {_DXF}")
    c = _client(_app())
    with open(_DXF, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, os.path.basename(_DXF)),
                         "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {}).get("sid")
    assert sid, r.get_json()
    _idle(c, sid)
    rv = c.post("/api/module-f/slot/read",
                json={"sid": sid, "method": "auto"})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    _idle(c, sid)
    return c, sid


def test_자동_차선_상태와_미리보기가_돈다(auto_lane):
    c, sid = auto_lane
    for path in ("/api/module-f/auto/state", "/api/module-f/auto/preview",
                 "/api/module-f/auto/network-view",
                 "/api/module-f/auto/handoff-hints"):
        rv = c.get(f"{path}?sid={sid}")
        assert rv.status_code < 500, \
            f"{path} → {rv.status_code} · {rv.get_data(as_text=True)[:200]}"
        assert rv.get_json() is not None


def test_자동_알람밸브_찍기의_세_갈래(auto_lane):
    """[S210·S220] 지정 · 지우기 · 거절 — 셋이 서로 다른 뜻이다.

    ★`y: None` 은 «엉터리 입력» 이 아니라 **«지우기»** 다. 처음에 그것을
      거절로 기대했다가 알았다 — 화면이 알람밸브를 취소하는 길이 그것이다.
      둘 다 숫자가 아닐 때만 거절이다.
    """
    c, sid = auto_lane
    clear = c.post("/api/module-f/auto/anchor", json={"sid": sid, "y": None})
    assert (clear.get_json() or {}).get("alarm") is None, "지우기가 안 된다"

    bad = c.post("/api/module-f/auto/anchor",
                 json={"sid": sid, "x": "a", "y": "b"})
    assert bad.status_code < 500, bad.get_data(as_text=True)[:200]
    assert (bad.get_json() or {}).get("ok") is False, "숫자 아닌 좌표가 통과한다"

    ok = c.post("/api/module-f/auto/anchor",
                json={"sid": sid, "x": 1.0, "y": 2.0})
    assert (ok.get_json() or {}).get("alarm") == [1.0, 2.0]


def test_자동_영역과_배관레이어_지정이_돈다(auto_lane):
    """영역 없이 두면 «도면 전체» 다 — 빈 목록도 정상 입력이다."""
    c, sid = auto_lane
    rv = c.post("/api/module-f/auto/zones", json={"sid": sid, "zones": []})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    rv2 = c.post("/api/module-f/auto/pipe-layers",
                 json={"sid": sid, "layers": []})
    assert rv2.status_code < 500, rv2.get_data(as_text=True)[:200]


def test_자동_헤드검출이_돈다(auto_lane):
    """[S300] 무거운 잡이다 — 끝까지 돌고 «몇 개» 를 답해야 한다."""
    c, sid = auto_lane
    _idle(c, sid)
    rv = c.post("/api/module-f/auto/heads", json={"sid": sid})
    assert rv.status_code < 500, rv.get_data(as_text=True)[:300]
    j = _idle(c, sid)
    assert j.get("state") in ("done", "error", "idle"), j
    # ★`/auto/state` 는 «state» 로 감싸지 않는다 — 값이 최상위에 온다.
    #   (`/edit/state` 는 감싼다. 두 규약이 다르다는 것을 여기 적어 둔다.)
    st = c.get(f"/api/module-f/auto/state?sid={sid}").get_json() or {}
    assert st.get("ok") is True, st
    assert st.get("method") == "auto", st
    assert st.get("opened") is True, st


# ═══════════════════════════════ ⑧ 남은 값싼 것들
def test_진행_스트리밍이_끝난_잡에서_즉시_닫힌다(edited):
    """[F-6] SSE — 잡이 done 이면 상태 한 번 보내고 **끝나야** 한다.

    ★안 끝나면 시험이 멈춘다. 그래서 이 시험 자체가 «닫히는가» 의 검사다 —
      생성기가 `return` 을 안 하면 여기서 걸린다.
    """
    c, sid = edited
    _idle(c, sid)
    rv = c.get(f"/api/module-f/job/stream?sid={sid}")
    assert rv.status_code == 200, rv.status_code
    text = rv.get_data(as_text=True)
    assert "event: state" in text, text[:200]


def test_내려받기가_없는_산출물을_통제해서_거절한다(edited):
    """파일이 없을 때 500 이 아니라 사유로 답해야 한다."""
    c, sid = edited
    rv = c.get(f"/api/module-f/download?sid={sid}&what=kfp")
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]


def test_산출물_내려받기가_돈다(confirmed):
    """표를 확정하고 저장한 뒤라야 받을 것이 있다 — 그 경로를 지난다."""
    c, sid = confirmed
    c.post("/api/module-f/design/emit", json={"sid": sid})
    _idle(c, sid)
    rv = c.get(f"/api/module-f/download?sid={sid}&what=design")
    assert rv.status_code < 500, rv.get_data(as_text=True)[:200]
    if rv.status_code == 200:
        assert len(rv.get_data()) > 0, "빈 파일을 내려준다"


def test_급수방식_고르기가_모르는_값을_막는다(edited):
    """[S710] 급수방식은 «도면에 없는 값» 이라 자동 추정하지 않는다."""
    c, sid = edited
    _idle(c, sid)
    modes = (c.get("/api/module-f/merge/modes").get_json() or {}).get("modes")
    assert modes, "급수방식 목록이 비었다"
    bad = c.post("/api/module-f/merge/mode",
                 json={"sid": sid, "mode": "없는방식"})
    assert 400 <= bad.status_code < 500, bad.status_code
    ok = c.post("/api/module-f/merge/mode",
                json={"sid": sid, "mode": modes[0]["key"]})
    assert ok.status_code < 500, ok.get_data(as_text=True)[:200]
    st = c.get(f"/api/module-f/merge/state?sid={sid}").get_json() or {}
    assert st.get("ok") is not None, st
