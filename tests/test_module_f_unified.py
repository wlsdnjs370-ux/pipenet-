# -*- coding: utf-8 -*-
"""[F-10] 세 차선 통합 — 흐름과 화면만 합치고 «코드는 합치지 않는다».

2026-08-27 상무 시연에서 28분 내내 같은 흐름 하나만 반복해서 요구했다
(전사 06:48 · 07:44 · 17:15 · 22:54 · 25:38). 사용자 머릿속에 «차선» 이라는
개념이 없다 — 업로드 시점에는 이 도면이 자동으로 될지 사람도 모르므로
「어떻게 추출할까요」는 **답할 수 없는 질문**이었다.

그래서 없애는 것은 질문 하나다. 엔진 두 경로는 그대로 산다 — A 의 자동 추출과
E 의 물길 판정은 「이어져 있다」의 정의가 달라, 그래프를 섞으면 G-BLOCKED B4
실측(헤드 물닿음 0 · 노드 2)이 재현된다.

이 파일은 그 «합치지 않음» 을 지키는 시험이다.
"""
from __future__ import annotations

import os
import re

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _screen() -> str:
    """화면 소스 한 덩이 — 마크업 + CSS + JS.

    자산이 세 파일로 나뉘어 있다. 이 시험들이 보는 것은 «무엇을 하는가» 지
    «어느 파일에 있는가» 가 아니므로 합쳐 읽는다.
    """
    parts = [open(os.path.join(_ROOT, "templates", "module_f.html"),
                  encoding="utf-8").read()]
    for rel in (("static", "module_f.css"), ("static", "module_f.js")):
        p = os.path.join(_ROOT, *rel)
        if os.path.isfile(p):
            parts.append(open(p, encoding="utf-8").read())
    return "\n".join(parts)


# ═══════════════════════════════════════════ F-10a. 방식 질문 제거
def test_방식_질문이_없다():
    """D-F10-1 — 업로드부터 손질까지 질문 0."""
    html = _screen()
    assert 'id="panel-method"' not in html
    for bid in ("mth-auto", "mth-mixed", "mth-manual", "mth-conf",
                "mth-cancel", "mth-recon", "mth-file"):
        assert f'id="{bid}"' not in html, f"#{bid} 가 남아 있다"
    # 「고르기 전에는 읽기를 시작할 수도 없다」던 옛 주석도 사실이 아니다.
    assert "고르기 전에는 읽기를 시작할 수도 없다" not in html


def test_흐름이_스스로_갈린다():
    """정찰 결과가 답한다 — 사람에게 되묻지 않는다."""
    html = _screen()
    i = html.index("function reconReady()")
    seg = html[i:i + 1200]
    # 갈림의 세 이유가 전부 «사유 문장» 을 들고 있다.
    assert "자동 인식이 실패했습니다" in seg
    assert "배관 레이어를 찾지 못했습니다" in seg
    assert "높음(≥0.9)» 헤드가 없어" in seg
    j = html.index("async function autoStart()")
    body = html[j:j + 1200]
    assert "reconReady()" in body
    assert "startNote(gate.why, true)" in body, "왜 그리 갔는지 안 적는다"
    # 갈림에 confirm/prompt 가 끼면 그것이 곧 질문이다.
    assert "confirm(" not in body and "prompt(" not in body


def test_기본_기준을_저절로_낮추지_않는다():
    """D-F8-4 는 그대로다 — 낮추는 것은 사람이 고급에서 한다.

    LH306 은 높음 띠가 0 이라(0/42) 0.9 로는 헤드가 하나도 안 찍히고, 그
    스펙으로 조립하면 엔진이 `KeyError: 'heads'` 로 죽는다(실측). 그래서 «막고
    사유를 적는» 쪽을 골랐다 — 기준을 프로그램이 낮추면 사람이 모르는 사이에
    낮은 신뢰도 후보가 산출에 들어간다.
    """
    html = _screen()
    i = html.index("function reconReady()")
    seg = html[i:i + 1200]
    assert "reconPick(0.9)" in seg, "기본 기준으로 재지 않는다"
    assert 'sel.value = "0.9"' in html
    # 자동으로 기준을 갈아끼우는 코드가 없어야 한다.
    j = html.index("async function autoStart()")
    assert re.search(r'\$\("adv-conf"\)\.value\s*=', html[j:j + 1200]) is None


def test_확정_지점은_손질이고_되돌릴_수_있다():
    """D-F10-3 — «확정은 사람» 은 유지, 자리만 손질로 옮겼다."""
    html = _screen()
    i = html.index("async function adoptRun(")
    seg = html[i:i + 2200]
    assert "/api/module-f/pick/commit" in seg
    assert "「찍기」로 내려가 고칠 수 있습니다" in seg


# ═══════════════════════════════════════════ 보존 — 합치지 않는다
def test_엔드포인트를_하나도_안_없앤다():
    """지시서 §3 — 없애는 것은 `panel-method` 화면 조각뿐이다."""
    import importlib

    srv = importlib.import_module("대조 서버")
    rules = {str(r.rule) for r in srv.app.url_map.iter_rules()}
    must = [
        "/api/module-f/auto/state", "/api/module-f/auto/anchor",
        "/api/module-f/auto/heads", "/api/module-f/auto/network",
        "/api/module-f/auto/run", "/api/module-f/auto/preview",
        "/api/module-f/auto/handoff", "/api/module-f/auto/pipe-layers",
        "/api/module-f/pick/adopt", "/api/module-f/pick/suggest",
        "/api/module-f/pick/commit", "/api/module-f/slot/read",
    ]
    missing = [p for p in must if p not in rules]
    assert not missing, f"엔드포인트가 사라졌다: {missing}"


def test_자동_차선_입구가_남아_있다():
    """D-F10-2 — 화면에서 «질문» 이 아니라 «설정» 이 됐을 뿐이다."""
    html = _screen()
    assert 'id="adv-auto"' in html
    assert 'readSlot("auto")' in html
    # 특허 실시예의 자동 화면 자체는 그대로다.
    for aid in ("au-anchor", "au-heads", "au-network", "au-run"):
        assert f'id="{aid}"' in html, f"#{aid} 가 없다"


def test_그래프_이식_코드가_없다():
    """지시서 §3 — A 선정 결과를 E board 로 옮기는 길을 만들지 않는다.

    기본 흐름은 «찍기 클릭» 으로만 board 에 들어간다(D-F10-6). 채택도 그
    길이다 — `adopt_heads` 가 `PickSession.click` 을 태운다.
    """
    from routes.module_f import adopt

    src = open(adopt.__file__, encoding="utf-8").read()
    assert ".click(" in src, "클릭 경로를 안 탄다"
    for banned in ("board.disks[", "board.mat[", "board.edges.append"):
        assert banned not in src, f"board 에 직접 쓴다: {banned}"


def test_흐름이_수동_경로를_그대로_쓴다():
    """새 흐름을 만들지 않는다 — 기존 찍기·손질 라우트만 잇는다."""
    html = _screen()
    i = html.index("async function autoStart()")
    seg = html[i:i + 1200]
    assert 'method: "manual"' in seg
    assert 'S.method = "manual"' in seg


# ═══════════════════════════════════════════ F-10b. 알람밸브 원클릭
#
# 시험 도면은 «대명동 단위세대» 다. LH306 은 헤드가 망에 안 닿아(실측: 어느
# 배관 자리에 급수원을 놓아도 「닿는 헤드가 없습니다」) 최불리 자체가 성립하지
# 않는다 — F-10b 가 재려는 것과 무관한 이유로 실패하는 도면이다.
_DXF = os.path.join(_ROOT, "routes", "제출용[최종]",
                    "1. 입력도면 대명동 단위세대 평면도.dxf")
MODE_VALVE = "알람밸브위치"
MODE_SOURCE = "급수시작위치"
ANCHOR_MAX_D = 2000.0


def _client(tmp_path=None):
    """앱 + 로그인. tmp_path 를 주면 «쓰기 루트» 를 그리로 돌린다.

    ★찍은스펙·표시캐시는 사용자의 작업 폴더(`docs/import`)에 쌓인다. 시험이
      거기에 쓰면 사용자가 손으로 찍어 둔 저장본을 덮는다 — 계측 도구가 같은
      이유로 임시 폴더를 쓴다(`scripts/_measure_module_f_lanes.py`).
    """
    import importlib

    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    if tmp_path is not None:
        from routes.module_f.common import _boot
        _boot()
        from services.cad_import.pipeline import disp_cache, handoff
        work = str(tmp_path)
        handoff.import_write_root = lambda: work
        handoff.OUT_DIR = handoff.pick_out_dir()
        disp_cache._DISP_CACHE_DIR = work
        os.makedirs(handoff.pick_out_dir(), exist_ok=True)
        os.makedirs(handoff.default_edits_dir(), exist_ok=True)
    c = srv.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def _wait(c, sid, limit=3000):
    import time

    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _build_edit(c, dxf, conf=0.75):
    """도면 하나를 손질 상태까지 — F-10a 의 기본 흐름과 같은 순서다."""
    with open(dxf, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, os.path.basename(dxf)), "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {})["sid"]
    _wait(c, sid)
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True, "heads": {"conf_min": conf}})
    assert _wait(c, sid)["state"] == "done"
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    assert _wait(c, sid)["state"] == "done"
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    return sid, st


def _pipe_points(state):
    """배관 선분의 끝점들 — 클릭이 붙을 자리.

    `body_groups[].segs` 는 «평평한 정수 mm 배열» 이다(x1,y1,x2,y2 반복) —
    도면 한 장이 수백 m 인데 0.1mm 를 실어 나를 이유가 없다는 화면 규약.

    ★큰 덩이부터 준다. 조각난 도면에서 아무 점이나 고르면 급수원이 헤드에
      안 닿아 「닿는 헤드가 없습니다」로 끝난다 — 시험하려는 것과 무관한 실패다.
    """
    out = []
    groups = sorted((g.get("segs") or [] for g in state.get("body_groups") or []),
                    key=len, reverse=True)
    for s in groups:
        for i in range(0, len(s) - 3, 4):
            out.append((float(s[i]), float(s[i + 1])))
            out.append((float(s[i + 2]), float(s[i + 3])))
    if not out:
        raise AssertionError("배관 선분이 없다")
    return out


def _anchor_points(state):
    """급수원을 놓아도 «헤드에 닿는» 배관 자리들.

    두 함정을 다 피해야 한다(둘 다 LH306 실측):
      · 배관에서 아무 점이나 고르면 그 조각에 헤드가 없어 「닿는 헤드가
        없습니다」로 끝난다.
      · 그렇다고 헤드 좌표를 그대로 찍으면 거기엔 배관이 없어 클릭이 안 붙는다.
    그래서 «헤드에 가장 가까운 배관 끝점» 을 쓴다 — 배관 위이면서 그 헤드가
    붙은 망이다.
    """
    pipe = _pipe_points(state)
    heads = [(float(h[0]), float(h[1])) for h in (state.get("heads") or [])]
    if not heads:
        return pipe
    out, seen = [], set()
    for hx, hy in heads:
        p = min(pipe, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out or pipe


def _a_pipe_point(state):
    return _anchor_points(state)[0]


def test_원클릭이_수동_두픽과_같은_답을_낸다(tmp_path):
    """★F-10b 의 핵심 — 클릭 «경로» 를 정말로 타는지의 증명.

    같은 좌표에 밸브·급수원을 수동으로 두 번 찍고 `/edit/worst` 를 부른 결과와,
    `/edit/anchor-click` 한 번의 결과가 완전히 같아야 한다. 다르면 원클릭이
    어딘가에서 «다른 길» 로 값을 만들고 있다는 뜻이다(D-F10-6 위반).
    """
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)

    # ── 수동 두 픽 + worst
    sid_a, st = _build_edit(c, _DXF, conf=0.9)
    x, y = _a_pipe_point(st)
    for mode in (MODE_VALVE, MODE_SOURCE):
        c.post("/api/module-f/edit/mode", json={"sid": sid_a, "mode": mode})
        c.post("/api/module-f/edit/click",
               json={"sid": sid_a, "x": x, "y": y, "max_d": ANCHOR_MAX_D})
    ra = c.post("/api/module-f/edit/worst", json={"sid": sid_a, "k": 30})
    assert ra.status_code == 200, ra.get_json()
    manual = ra.get_json()["summary"]

    # ── 원클릭 한 번 (같은 좌표)
    sid_b, _ = _build_edit(c, _DXF, conf=0.9)
    rb = c.post("/api/module-f/edit/anchor-click",
                json={"sid": sid_b, "x": x, "y": y})
    assert rb.status_code == 200, rb.get_json()
    job = _wait(c, sid_b)
    assert job["state"] == "done", job.get("error")
    # 잡 «보기»(/job)는 진행만 싣는다 — 결과는 따로 청한다(기존 규약).
    res = (c.get(f"/api/module-f/convert/result?sid={sid_b}")
           .get_json() or {}).get("result") or {}
    one = res.get("summary")
    assert one is not None, res

    for key in ("k", "reachable", "far_m", "near_m", "span_m", "total_m",
                "max_load", "source", "candidates", "anchor_path_m",
                "anchor_path_nodes", "path_edges"):
        assert manual[key] == one[key], (
            f"«{key}» 가 다르다 — 수동 {manual[key]} vs 원클릭 {one[key]}")


def test_손질_기본_모드가_원클릭이다():
    """수용 기준 — 손질에 들어오면 첫 동작이 알람밸브 한 번이다.

    ★원클릭은 «서버 모드» 가 아니라 화면 모드다. 서버의 손질 모드는 이음·삭제·
      급수시작위치·알람밸브위치 넷 그대로이고, 원클릭은 그중 둘을 한 번에 놓는
      행동이다 — 엔진 계약을 늘리지 않는다.
    """
    html = _screen()
    assert 'data-mode="원클릭"' in html
    assert 'id="ed-anchor-note"' in html
    assert "알람밸브를 클릭하면 가장 불리한 배관망이 표시됩니다" in html
    i = html.index("async function loadEdit()")
    assert "setUiMode(ONECLICK)" in html[i:i + 700], "기본 모드가 아니다"
    # 화면 모드를 서버로 보내면 「모르는 손질 모드입니다」로 튕긴다.
    j = html.index('if (mode === ONECLICK)')
    assert "return" in html[j:j + 160]


def test_원클릭은_클릭_경로로만_넣는다():
    """D-F10-6 — board 에 직접 쓰는 코드가 없어야 한다."""
    import inspect

    from routes.module_f import api_edit

    src = inspect.getsource(api_edit)
    i = src.index("def module_f_edit_anchor_click")
    seg = src[i:i + 4000]
    assert "es.set_mode(" in seg and "es.click(" in seg
    for banned in ("board.sources.append", "board.valves.append",
                   "b.sources.append", "b.valves.append",
                   "b.sources =", "b.valves ="):
        assert banned not in seg, f"board 에 직접 쓴다: {banned}"


def test_원클릭_뒤_되돌리면_두_픽이_풀린다(tmp_path):
    """수용 기준 — undo 두 번이면 두 픽이 순서대로 되돌아간다."""
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)
    sid, st = _build_edit(c, _DXF, conf=0.9)
    x, y = _a_pipe_point(st)
    c.post("/api/module-f/edit/anchor-click", json={"sid": sid, "x": x, "y": y})
    assert _wait(c, sid)["state"] == "done"

    def picks():
        s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        return len(s.get("sources") or []), len(s.get("valves") or [])

    assert picks() == (1, 1), picks()
    c.post("/api/module-f/edit/undo", json={"sid": sid})
    assert picks() == (0, 1), "급수가 먼저 풀려야 한다(나중에 찍었으므로)"
    c.post("/api/module-f/edit/undo", json={"sid": sid})
    assert picks() == (0, 0), "밸브도 풀려야 한다"


def test_원클릭이_기존_픽을_갈아끼운다(tmp_path):
    """수용 기준 — 이미 있던 세션에서 다시 부르면 자리를 옮긴다."""
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)
    sid, st = _build_edit(c, _DXF, conf=0.9)
    pts = _anchor_points(st)
    p1 = pts[0]
    p2 = next((p for p in pts
               if abs(p[0] - p1[0]) + abs(p[1] - p1[1]) > 1.0), None)
    assert p2 is not None, "서로 다른 두 점이 없다"

    for (x, y) in (p1, p2):
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": x, "y": y})
        j = _wait(c, sid)
        assert j["state"] == "done", j.get("error")
    s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    # 두 번 불러도 «하나씩» 이어야 한다 — 쌓이면 급수원이 여럿이 된다.
    assert len(s["sources"]) == 1, s["sources"]
    assert len(s["valves"]) == 1, s["valves"]


# ═══════════════════════════════════════════ 정찰이 깨져도 흐름은 산다
def test_정찰이_실패해도_찍기는_열린다():
    """수용 기준 — 모듈 A 가 아예 안 되는 도면에서도 «묻지 않고» 찍기로.

    정찰이 실패하면 `run_recon` 이 사유를 담은 dict 를 돌려주고(잡을 죽이지
    않는다), 화면은 그 상태를 «자동으로 시작할 수 없음» 으로 읽는다. 여기서는
    실패 기록이 화면이 읽을 수 있는 모양으로 나오는지까지만 본다 — 갈림 자체는
    `test_흐름이_스스로_갈린다` 가 지킨다.
    """
    from routes.module_f.recon import recon_view

    assert recon_view(None)["state"] == "none"
    bad = recon_view({"error": "ModuleNotFoundError: remote30_prototype"})
    assert bad["state"] == "error"
    assert "remote30_prototype" in bad["error"], "사유를 안 넘긴다"
    ok = recon_view({"heads": [{"x": 0, "y": 0, "conf": 0.9}],
                     "bands": {"높음(≥0.9)": 1}, "bundles": {"PIPE": 2}})
    assert ok["state"] == "ok" and ok["n"] == 1


def test_정찰_실패가_열기를_죽이지_않는다():
    """찍기는 정찰과 무관하게 서야 한다 — 실패는 «올려» 부르는 쪽이 감싼다."""
    import inspect

    from routes.module_f import api_open

    src = inspect.getsource(api_open)
    i = src.index("run_recon")
    seg = src[max(0, i - 600):i + 600]
    assert "except" in seg, "정찰 실패를 감싸지 않는다 — 열기가 같이 죽는다"


def test_모듈_A_를_막아도_찍기까지_간다(monkeypatch, tmp_path):
    """수용 기준 — 「정찰을 일부러 깨뜨리면 질문 없이 찍기 화면 + 사유 배너」.

    모듈 A 의 인식을 통째로 못 쓰게 만든 뒤 실제로 도면을 열어, ① 열기 자체는
    성공하고 ② 찍기판이 서고 ③ 정찰이 «사유를 가진 실패» 로 보고되는지 본다.
    화면의 갈림은 이 세 가지에만 기대므로, 여기까지면 폴백이 성립한다.
    """
    import importlib
    import os as _os
    import time

    _os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = os.path.join(_ROOT, "samples", "dxf", "분기티.dxf")
    if not os.path.isfile(dxf):
        import pytest
        pytest.skip("표본 도면 없음")

    from routes.module_f import recon as recon_mod

    def boom(*a, **kw):
        raise ModuleNotFoundError("remote30_prototype (일부러 막음)")

    monkeypatch.setattr(recon_mod, "run_recon", boom)

    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, "분기티.dxf"), "kind": "plan"},
                       content_type="multipart/form-data")
        body = r.get_json() or {}
        assert r.status_code == 200 and body.get("sid"), body
        sid = body["sid"]
        for _ in range(600):
            j = c.get(f"/api/module-f/job?sid={sid}").get_json()
            if j.get("state") in ("done", "error", "idle"):
                break
            time.sleep(0.1)
        # ① 열기는 살아 있다  ② 찍기판이 섰다
        assert j.get("state") == "done", j.get("error")
        w = c.get(f"/api/module-f/world?sid={sid}")
        assert w.status_code == 200, w.get_json()
        # ③ 정찰은 «사유를 가진 실패» 로 보고된다 → 화면이 찍기로 가른다
        rec = (c.get(f"/api/module-f/recon?sid={sid}").get_json()
               or {}).get("recon") or {}
        assert rec.get("state") == "error", rec
        assert "일부러 막음" in str(rec.get("error")), rec
