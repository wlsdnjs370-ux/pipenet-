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
    # [F-11a · D-F11-2] 셋째 사유는 이제 «지배 띠 규칙» 의 문장을 그대로 쓴다.
    #   예전에는 「높음(≥0.9) 헤드가 없어」라고 절대 임계를 못 박았는데, 그
    #   임계 자체가 도면 분포로 바뀌었다(recon.dominant_band).
    assert "a.why" in seg and "고급에서 채택 기준을 낮춰" in seg
    j = html.index("async function autoStart()")
    body = html[j:j + 1200]
    assert "reconReady()" in body
    assert "startNote(gate.why, true)" in body, "왜 그리 갔는지 안 적는다"
    # 갈림에 confirm/prompt 가 끼면 그것이 곧 질문이다.
    assert "confirm(" not in body and "prompt(" not in body


def test_기준을_프로그램이_정하되_반드시_말한다():
    """[D-F11-2 가 D-F8-4 를 개정했다] 기준을 도면 분포가 정한다.

    ★예전 규약은 「기준을 프로그램이 낮추지 않는다」였다. 그 이유는 「사람이
      모르는 사이에 낮은 신뢰도 후보가 들어간다」였는데, 절대 임계 0.9 는
      A 의 신뢰도가 사실상 이진값이라 도면마다 뒤집혀 **퇴화**했다
      (B1F 72/3,338 → 최불리 2개 · LH306 0/42 → 조립 불가).

      그래서 결정이 바뀌었다: 프로그램이 정하되 **반드시 말한다.** 「모르는
      사이에」가 사라지면 원래 걱정도 사라진다. 규칙은 결정적이고, 발동한
      규칙이 카드와 배너에 적히고, 사람이 고르면 사람이 이긴다.
    """
    html = _screen()
    i = html.index("function reconReady()")
    seg = html[i:i + 1400]
    assert "reconPick(confMin())" in seg, "아직 절대 임계로 판단한다"
    # 기본 임계는 «서버의 규칙» 에서 온다.
    j = html.index("function confMin()")
    body = html[j:j + 700]
    assert "S.recon" in body and "adopt" in body
    # 화면이 제멋대로 임계 칸을 갈아끼우지 않는다 — 사람이 고른 값을 덮으면
    # 그것이 곧 「모르는 사이에」다.
    k = html.index("async function autoStart()")
    assert re.search(r'\$\("adv-conf"\)\.value\s*=', html[k:k + 1200]) is None


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

    같은 좌표를 수동으로 찍고 `/edit/worst` 를 부른 결과와, `/edit/anchor-click`
    한 번의 결과가 완전히 같아야 한다. 다르면 원클릭이 어딘가에서 «다른 길» 로
    값을 만들고 있다는 뜻이다(D-F10-6 위반).

    ★수동 픽도 이제 **한 번** 이다. 알람밸브가 접속점을 겸하므로(B1 해소),
      같은 자리를 두 모드로 두 번 찍으면 토글이 되어 «지운» 것이 된다.
    """
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)

    # ── 수동 두 픽 + worst
    sid_a, st = _build_edit(c, _DXF, conf=0.9)
    x, y = _a_pipe_point(st)
    c.post("/api/module-f/edit/mode", json={"sid": sid_a, "mode": MODE_VALVE})
    c.post("/api/module-f/edit/click",
           json={"sid": sid_a, "x": x, "y": y, "max_d": ANCHOR_MAX_D})
    # ★픽 한 번이 접속점까지 놓는다 — 이것이 B1 해소의 알맹이다.
    sa = c.get(f"/api/module-f/edit/state?sid={sid_a}").get_json()["state"]
    assert len(sa.get("valves") or []) == 1, sa.get("valves")
    assert (sa.get("sources") or []) == (sa.get("valves") or []),         f"알람밸브와 접속점이 어긋났다 — {sa.get('sources')} vs {sa.get('valves')}"
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
                "max_load", "source", "candidates", "worst_path_m",
                "worst_path_nodes", "path_edges"):
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


def test_원클릭_뒤_되돌리면_한_번에_풀린다(tmp_path):
    """★한 번의 클릭은 한 번의 되돌리기로 풀려야 한다.

    종전에는 밸브·급수를 따로 찍었으므로 undo 를 «두 번» 해야 했다. 사람이
    보기엔 한 번 누른 것인데 두 번 눌러야 돌아가는 것은 그 자체로 어긋남이고,
    한 번만 누르면 «접속점 없는 알람밸브» 라는 중간 상태가 남았다.
    """
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
    assert picks() == (0, 0),         "한 번 찍은 것이 한 번에 안 풀렸다 — 중간 상태가 남는다"


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


# ═══════════════════════════════════════════ F-10c. 시각 위계
def test_배경_도면이_사라지지_않는다():
    """전사 08:57 「도면은 그대로 있잖아」 · 17:53 「밑에 배경도면 보이잖아」.

    위계는 «지우기» 가 아니라 «농도» 다. 배경이 사라지면 어디서 뽑힌 망인지
    안 보여 결과가 옳은지 판단할 수가 없다.
    """
    html = _screen()
    assert 'id="ed-bg"' in html and "checked" in html
    assert "EDIT_BG_ALPHA" in html
    i = html.index("const EDIT_BG_ALPHA")
    # 아주 흐리되 0 은 아니다 — 0 이면 «사라진» 것이다.
    m = re.search(r"const EDIT_BG_ALPHA\s*=\s*([0-9.]+)", html[i:i + 80])
    assert m and 0.0 < float(m.group(1)) < 0.2, html[i:i + 60]
    # 비corridor 배관망의 기본은 «감추기» 가 아니라 «흐리게» 다.
    j = html.index('id="ed-worst-view"')
    seg = html[j:j + 400]
    assert 'value="dim" selected' in seg, "기본이 흐리기가 아니다"


def test_펄스는_몇_번_하고_멈춘다():
    """전사 06:57 «반짝반짝» 의 의도는 강조지 점멸 «지속» 이 아니다.

    계속 깜빡이면 눈이 피로하고, 캔버스를 매 프레임 다시 그리므로 큰 도면에서
    비용도 계속 든다.
    """
    html = _screen()
    i = html.index("function pulseAmt()")
    seg = html[i:i + 500]
    assert "pulseT0 = 0" in seg, "스스로 멈추지 않는다"
    m = re.search(r"const PULSE_MS\s*=\s*(\d+)", html)
    assert m and int(m.group(1)) <= 3000, "너무 오래 반짝인다"
    m2 = re.search(r"const PULSE_CYCLES\s*=\s*([0-9.]+)", html)
    assert m2 and float(m2.group(1)) <= 3.0, "2~3회를 넘는다"
    # 새 corridor 가 나왔을 때만 시작한다 — 두 길(원클릭·최불리 선정) 모두에서.
    assert html.count("startPulse()") >= 3


def test_위계_토글은_표시_전용이다():
    """수용 기준 — 토글로 왕복해도 세션 상태가 안 바뀐다.

    토글이 서버를 부르면 그것은 표시가 아니라 «상태 변경» 이다. 두 토글의
    onchange 가 `draw()` 하나만 부르는지 소스로 못 박는다.
    """
    html = _screen()
    for el in ("ed-worst-view", "ed-bg"):
        i = html.index(f'$("{el}").onchange')
        line = html[i:html.index("\n", i)]
        assert "draw()" in line, line
        assert "post(" not in line and "api(" not in line, line


def test_선정_헤드는_동그라미다():
    """전사 23:45 「헤드를 그냥 이렇게 역력하는 것보다 동그라미를 치는 게 더
    보기가 좋다」 — 선정 30개는 고리로, 앵커는 겹원으로."""
    html = _screen()
    i = html.index("for (const h of e.worst.heads)")
    seg = html[i:i + 400]
    assert "ctx.arc(" in seg and "ctx.stroke()" in seg
    j = html.index("if (e.worst.worst_head)")
    assert "#ff3b3b" in html[j:j + 400], "앵커가 따로 강조되지 않는다"


# ═══════════════════════════════════════════ F-10d. 결과 위 수정
def test_수정을_세고_다시_계산하면_0으로_돌아온다(tmp_path):
    """수용 기준 — 수정 3건 후 배지가 3건, 다시 계산 후 0.

    ★수정마다 최불리를 다시 돌리지 않는다(D-F10-5). 검출이 실측 ~18초라
      클릭 하나가 18초짜리가 되어 버린다.
    """
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)
    sid, st = _build_edit(c, _DXF, conf=0.9)
    x, y = _a_pipe_point(st)
    c.post("/api/module-f/edit/anchor-click", json={"sid": sid, "x": x, "y": y})
    assert _wait(c, sid)["state"] == "done"

    def state():
        return c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]

    s0 = state()
    assert s0["edits_since_worst"] == 0, "계산 직후인데 수정이 쌓여 있다"
    assert s0.get("worst"), "최불리가 없다"

    # 삭제 모드로 세 번 두드린다 — board 를 바꾸는 클릭이면 세어야 한다.
    c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": "삭제"})
    pts = _pipe_points(st)
    hit = 0
    for (px, py) in pts:
        if hit >= 3:
            break
        r = c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": px, "y": py, "max_d": 300.0})
        if (r.get_json() or {}).get("report"):
            hit += 1
    assert hit == 3, f"board 를 바꾸는 클릭이 {hit}번밖에 안 됐다"
    s1 = state()
    assert s1["edits_since_worst"] == 3, s1["edits_since_worst"]
    # 낡은 corridor 를 «진짜처럼» 남기지 않는다 — 절점 번호가 어긋난다.
    assert not s1.get("worst"), "낡은 최불리가 그대로 남았다"

    # 다시 계산 — 픽은 그대로다(급수원·밸브를 다시 안 찍는다).
    r = c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
    assert r.status_code == 200, r.get_json()
    s2 = state()
    assert s2["edits_since_worst"] == 0, s2["edits_since_worst"]
    assert s2.get("worst"), "다시 계산했는데 최불리가 없다"
    assert len(s2["sources"]) == 1 and len(s2["valves"]) == 1, "픽이 사라졌다"


def test_되돌리기도_수정으로_센다(tmp_path):
    """되돌리기도 망을 바꾼다 — 세지 않으면 배지가 거짓말을 한다."""
    import pytest

    if not os.path.isfile(_DXF):
        pytest.skip("표본 도면 없음")
    c = _client(tmp_path)
    sid, st = _build_edit(c, _DXF, conf=0.9)
    x, y = _a_pipe_point(st)
    c.post("/api/module-f/edit/anchor-click", json={"sid": sid, "x": x, "y": y})
    assert _wait(c, sid)["state"] == "done"
    c.post("/api/module-f/edit/undo", json={"sid": sid})
    s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    assert s["edits_since_worst"] >= 1, s["edits_since_worst"]


def test_다시_계산은_자동이_아니다():
    """D-F10-5 — 수정마다 worst 를 다시 돌리는 코드가 없어야 한다."""
    html = _screen()
    # 클릭 처리에서 곧바로 최불리를 부르면 그것이 자동 재실행이다.
    i = html.index("async function editClick(")
    seg = html[i:i + 1200]
    assert "/edit/worst" not in seg, "클릭이 최불리를 자동으로 다시 돌린다"
    # 다시 계산은 «사람이 누르는 단추» 다.
    assert 'id="ed-recalc"' in html
    assert '$("ed-recalc").onclick' in html
    j = html.index('$("ed-recalc").onclick')
    assert "runWorst(" in html[j:j + 200]
    # 배지는 서버가 센 값을 그대로 쓴다.
    assert "edits_since_worst" in html
    # ★수정→다시 계산 왕복이 «화면 전환 없이» 되려면 셋이 한 패널에 있어야
    #   한다: 손질 모드(살리기·제거) · 다시 계산 단추 · 배지.
    i = html.index('id="panel-edit"')
    j = html.index("</section>", i)
    panel = html[i:j]
    for frag in ('data-mode="이음"', 'data-mode="삭제"',
                 'id="ed-recalc"', 'id="ed-edits"'):
        assert frag in panel, f"{frag} 가 손질 패널 밖에 있다"


# ═══════════════════════════════════════════ F-10e. 밑그림 — «평면» 에서
#
# 지시서는 밑그림을 아이소 «아래» 에 깔라고 했다. 지금 화면은 «평면에서 본다» 다.
#
# ★2026-09-03 정정 — 그 선택의 근거였던 「설계 좌표계는 board 의 변환이 아니다」
#   는 **사실이 아니었다**(BLOCKED §17 정정). 짝짓기를 `edge_ref` 로 잡아 잔차가
#   9.3% 로 나온 것이고, 믿을 수 있는 대응으로 다시 재면 0.018% 다 —
#   board → 설계 «평면» 은 전역 닮음으로 이어진다.
#   아이소 «아래» 가 안 되는 것은 여전히 맞지만 이유가 다르다: 아이소는
#   (x,y,z) 의 아핀이라(잔차 0.00) 평면 좌표만으로는 자리가 안 정해진다.
#   그러므로 이 시험이 지키는 것은 「평면 밑그림이 있다」이지 「아이소는 불가」가
#   아니다 — 아이소 밑그림은 표고를 갖고 같은 아핀을 타면 된다(미구현).
def test_밑그림은_평면에서_본다():
    """[F-10e] 밑그림 + 그 자리 수정을 평면 화면에서 만족시킨다."""
    html = _screen()
    assert 'id="dg-plan"' in html, "평면에서 보기 토글이 없다"
    assert "planUnderlayOn" in html
    # 밑그림은 손질 화면과 «같은» 배경을 쓴다 — 두 화면이 다른 그림을 보이면
    # 어느 쪽이 사실인지 알 수 없다.
    i = html.index("if (planUnderlayOn() && S.edit) {")
    seg = html[i:i + 400]
    assert "drawWorld(true, EDIT_BG_ALPHA)" in seg
    assert "drawEdit()" in seg


def test_평면_밑그림에서_그_자리_수정된다():
    """클릭은 손질과 «같은 경로» 다 — 새 길을 만들지 않는다(D-F10-6)."""
    html = _screen()
    i = html.index('S.stage === "design" && planUnderlayOn()')
    seg = html[i:i + 200]
    assert "editClick(" in seg, "설계 화면 클릭이 손질 경로를 안 탄다"
    # 다시 계산 → 표 확정 → 아이소 갱신이 한 단추다.
    j = html.index('$("dg-recalc").onclick')
    body = html[j:j + 900]
    assert "/api/module-f/edit/worst" in body
    assert '$("dg-build").click()' in body, "아이소가 안 갱신된다"


def test_아이소_좌표에_손대지_않았다():
    """수용 기준 — 밑그림 켬/끔은 표시 전용, 끔 상태는 종전 아이소와 같다.

    ★G16: 아이소에 그리는 corridor 좌표는 emit 에 쓰는 그 사본이다. 밑그림
      때문에 좌표를 손보면 미리보기가 거짓말이 된다.
    """
    html = _screen()
    # 토글은 화면만 다시 그린다 — 서버를 부르는 것은 «손질 상태 받아오기»
    # 하나뿐이고 그것도 읽기(GET)다.
    # ★창을 «다음 핸들러 앞» 에서 끊는다. 넉넉히 잡으면 옆 핸들러의 정당한
    #   호출이 딸려 들어와 없는 결함을 잡는다(실측으로 한 번 겪었다).
    j = html.index('$("dg-plan").onchange')
    end = html.index('document.querySelectorAll(".dgmode")', j)
    seg = html[j:end]
    assert "post(" not in seg, "토글이 서버 상태를 바꾼다"
    assert "draw()" in seg
    # 설계 미리보기는 여전히 엔진의 display_tables 결과를 그대로 쓴다.
    import inspect

    from routes.module_f import api_design

    src = inspect.getsource(api_design)
    assert "display_tables" in src
    assert "underlay" not in src, "설계 라우트에 밑그림 좌표가 섞였다"


def test_밑그림은_산출물을_안_건드린다():
    """표시 전용 증명 — 서버에 밑그림 상태를 저장하는 자리가 없다."""
    import inspect

    from routes.module_f import api_design, views

    for mod in (api_design, views):
        src = inspect.getsource(mod)
        for banned in ("dg_plan", "plan_underlay", "underlay"):
            assert banned not in src, f"{mod.__name__} 에 {banned} 가 있다"


# ═══════════════════════════════════════════ F-10f. 이상 표시
def test_이상_목록이_있고_0이면_완료를_말한다():
    """전사 27:41 「뭔가 좀 이상하면 표시를 해서 확인을 해서 수정을 하고」.

    목록이 0 이면 그것이 사람 검수의 «완료 신호» 다 — 「확인할 이상 없음」.
    """
    html = _screen()
    assert 'id="dg-issues"' in html and 'id="dg-issues-n"' in html
    assert "확인할 이상 없음" in html
    i = html.index("function renderIssues()")
    seg = html[i:i + 1200]
    assert "확인할 이상 없음" in seg, "0 일 때 완료를 안 말한다"
    # ★안 재고 «없다» 고 하면 완료 신호를 위조하는 것이다. 표가 없으면
    #   «아직 모른다» 여야 한다(저장소 규약: 정직한 진행 표시).
    assert "if (!S.design || !S.design.view)" in seg
    assert "표를 확정하면" in seg


def test_이상_목록의_합계가_요약과_같은_자료다():
    """수용 기준 — 같은 데이터의 두 얼굴.

    목록이 새 계산을 하면 요약과 어긋날 수 있다. 그래서 **이미 화면에 온
    자료** 만 쓴다: 관경 근거는 배관 행의 `src`, 부속·등가길이는 요약 수치
    그대로, 제외 사유는 `marks`, 유령은 채택 결과.
    """
    html = _screen()
    i = html.index("function collectIssues()")
    seg = html[i:i + 4200]
    # 관경 폴백은 요약의 nfpc_fallback 과 같은 판정을 쓴다.
    assert 'p.src === "nfpc_fallback"' in seg
    # [§18] 부속·등가길이는 «엔진이 센 그 자리» 의 목록을 그대로 쓴다.
    #   F 가 다시 판정하면 규칙이 두 벌이 되어 언젠가 갈린다.
    assert "S.design.tables.unresolved" in seg
    assert "kind_items" in seg and "length_items" in seg
    # 제외 사유는 F-5 의 marks 를 그대로 쓴다.
    for k in ("dry", "unattached", "unpicked"):
        assert f'"{k}"' in seg
    # 새로 서버를 부르지 않는다 — 부르면 그 순간 요약과 다른 시점이 된다.
    assert "api(" not in seg and "post(" not in seg


def test_자리를_모르는_항목은_클릭_대상이_아니다():
    """좌표가 없는 항목을 눌러도 엉뚱한 데로 가면 안 된다.

    [§18] 부속·등가길이는 이제 «어느 배관인지» 를 엔진에서 받으므로 대개
    좌표가 있다. 그래도 표에 없는 배관 id 가 오면 좌표가 비는데, 그때
    누르면 아무 일도 없어야 한다.
    """
    html = _screen()
    j = html.index("el.onclick = () => {", html.index("function renderIssues()"))
    assert "it.x !== null" in html[j:j + 300]


def test_미해결_목록은_엔진이_센_그_자리에서_나온다():
    """[§18] 목록과 개수가 어긋날 수 없다 — 같은 자리에서 나오기 때문이다.

    F 가 부속 판정을 다시 구현하면 규칙이 두 벌이 되고, 그 어긋남은
    «엉뚱한 배관을 미해결로 찍어 사람이 멀쩡한 값을 덮어쓰는» 형태로
    드러난다(표시 오류가 아니라 계산 오류다). 그래서 엔진이 세는 그 줄
    바로 옆에서 목록도 함께 쌓는다.
    """
    import inspect

    from services.cad_import.design import fitting

    src = inspect.getsource(fitting.build_fittings)
    # 세는 곳과 담는 곳이 붙어 있어야 한다.
    for count_line, item_line in (
        ("unresolved_kind += bad", "unresolved_kind_items.append"),
        ("unresolved_length += 1", "unresolved_length_items.append"),
    ):
        assert count_line in src and item_line in src
        i = src.index(count_line)
        assert item_line in src[i:i + 700], f"{item_line} 가 세는 자리에서 멀다"
    # 반환에 셋 다 실린다.
    for key in ("unresolved_kind_items", "unresolved_length_items",
                "unresolved_pairs"):
        assert f'"{key}"' in src


def test_직접_입력은_못_가린_자리에만_쓰인다():
    """[§18] ★가장 중요한 안전 성질 — 규칙이 낸 값은 안 바뀐다.

    사람이 한 자리를 채웠다고 산출 전체가 조용히 달라지면 안 된다. 그래서
    덮어쓰기 조회는 «판정 불가일 때» 안에서만 일어난다. 소스로 못 박는다 —
    조회가 그 밖으로 나가는 순간 «채우기» 가 «덮어쓰기» 로 바뀐다.
    """
    import inspect
    import re as _re

    from services.cad_import.design import fitting

    src = inspect.getsource(fitting.build_fittings)
    probe = "ov_kind.get("
    i = src.index(probe)
    # 그 줄 앞쪽 400자 안에 «미해결일 때» 라는 조건이 있어야 한다.
    before = src[max(0, i - 400):i]
    assert _re.search(r"if bad:", before), \
        f"{probe} 가 «판정 불가» 밖에서 불린다"


def test_등가길이_직접입력은_라이브러리를_못_덮는다():
    """등가길이 쪽 같은 성질 — 소스가 아니라 **행위**로 본다.

    종전에는 `build_fittings` 소스에서 `ov_eq.get(` 리터럴을 찾고 그 앞줄에
    조건이 있는지 봤다. 결정 로직을 `resolve_eq_len` 한 곳으로 모으자 그
    리터럴이 사라져 시험이 부러졌다 — 성질은 그대로인데 검사가 부러진 것이라,
    잡던 것을 실제로 잡게 고쳐 쓴다.
    """
    from services.cad_import.design.fitting import (
        load_equivalent_lengths, resolve_eq_len)
    lib = load_equivalent_lengths()

    # ① 라이브러리에 «있는» 자리는 사람이 무슨 값을 넣어도 안 바뀐다.
    have, why = resolve_eq_len("elbow", 100, lib=lib)
    assert have is not None and why == "라이브러리"
    same, why2 = resolve_eq_len("elbow", 100, lib=lib,
                                ov_eq={("elbow", 100): (99.0, "덮기 시도")})
    assert (same, why2) == (have, why), "채우기가 덮어쓰기로 바뀌었다"

    # ② 라이브러리에 «없는» 자리에만 사람이 채운 값이 들어간다.
    assert resolve_eq_len("alarm_valve", 15, lib=lib) == (None, None)
    got, note = resolve_eq_len("alarm_valve", 15, lib=lib,
                               ov_eq={("alarm_valve", 15): (6.5, "KFI")})
    assert (got, note) == (6.5, "KFI")

    # ③ ★못 구하면 0 이 아니라 None 이다. 0 은 「손실이 없다」는 주장이라,
    #    못 구한 것과 같은 값으로 두면 계산이 조용히 낙관적으로 틀어진다.
    assert resolve_eq_len("elbow", 9999, lib=lib) == (None, None)
    assert resolve_eq_len("elbow", None, lib=lib) == (None, None)


def test_직접_입력은_아는_종류만_고르게_한다():
    """자유 입력이면 라이브러리에 없는 이름이 들어와 등가길이가 다시 미해결이 된다.

    실측으로 겪었다 — 「엘베」라고 적었더니 부속 판정 불가는 3→2 로 줄면서
    등가길이 미해결이 0→1 로 늘었다. 그래서 서버가 고를 수 있는 목록을 준다.
    """
    import inspect

    from routes.module_f import api_design

    src = inspect.getsource(api_design)
    i = src.index("def module_f_design_fitting_override_get")
    seg = src[i:i + 2000]
    assert '"kinds"' in seg
    assert "ELBOW_45" in seg and "ELBOW_90" in seg and "TEE" in seg
    # 「직선 — 부속 없음」도 정답의 하나다(22.5° 미만).
    assert '"none"' in seg and "직선" in seg


def test_직접_입력은_산출물에도_남는다():
    """자동이 낸 값과 사람이 넣은 값을 같은 얼굴로 두지 않는다(모듈 A 방식)."""
    import inspect

    from services.cad_import.design import tables

    src = inspect.getsource(tables.build_design_tables)
    assert "직접 입력 — 부속 판정" in src
    assert "직접 입력 — 등가길이" in src
    # 사유도 함께 남아야 한다 — 「누가 왜 정했나」가 값과 같이 있어야 한다.
    fsrc = inspect.getsource(
        __import__("services.cad_import.design.fitting",
                   fromlist=["build_fittings"]).build_fittings)
    assert '"note"' in fsrc and "applied_overrides" in fsrc


def test_직접_입력_라우트가_엉터리_값을_막는다():
    """빈 칸·음수·숫자 아닌 값이 조용히 계산에 들어가면 안 된다."""
    import importlib

    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        from routes.module_f import jobs
        sess = jobs._new_session()
        sid = sess["id"]
        bad = [
            {"kind": [{"node": "N1", "pipe": "", "kind": "elbow"}]},
            {"eq_len": [{"kind": "elbow", "dia": "가나", "m": 1.0}]},
            {"eq_len": [{"kind": "elbow", "dia": 40, "m": -1}]},
            {"kind": [{"node": "N1", "pipe": "P1", "kind": "elbow",
                       "note": "x" * 201}]},
        ]
        for body in bad:
            r = c.post("/api/module-f/design/fitting-override",
                       json={"sid": sid, **body})
            assert r.status_code == 400, (body, r.get_json())
        # 옳은 값은 통과하고, «다시 확정하라» 고 말한다.
        r = c.post("/api/module-f/design/fitting-override", json={
            "sid": sid,
            "kind": [{"node": "N1", "pipe": "P1", "kind": "none",
                      "note": "현장 확인"}]})
        assert r.status_code == 200, r.get_json()
        assert "표 확정" in (r.get_json() or {}).get("message", "")


def test_화면이_그_자리에서_채우게_한다():
    """[§18] 항목 바로 아래에 채우는 칸 — 따로 떨어뜨리면 어느 자리 값인지 모른다."""
    html = _screen()
    assert 'id="dg-ov-save"' in html and 'id="dg-ov-n"' in html
    i = html.index("function renderIssues()")
    seg = html[i:i + 2600]
    assert "it.ov" in seg, "채울 수 있는 자리를 안 가린다"
    assert 'class="ovk"' in seg and 'class="ovm"' in seg
    assert 'class="ovn"' in seg, "사유 칸이 없다"
    # 종류는 서버가 준 목록에서만 고른다 — 자유 입력은 문제를 옮길 뿐이다.
    assert "S.fitKinds" in seg
    assert "loadFitKinds" in html


def test_저장하면_다시_확정하라고_말한다():
    """값이 바뀌는 일이라 표시만 고치고 끝내면 안 된다.

    저장은 세션에만 남는다 — 표를 다시 확정해야 산출에 들어간다. 화면이 그
    사실을 말하고, 실제로 다시 확정까지 이어 준다.
    """
    html = _screen()
    i = html.index('$("dg-ov-save").onclick')
    seg = html[i:i + 2000]
    assert "/api/module-f/design/fitting-override" in seg
    assert '$("dg-build").click()' in seg, "다시 확정으로 안 이어진다"
    # 안내 문구가 «못 가린 자리에만» 이라는 성질을 밝힌다.
    assert "못 가린 자리에만" in html
    assert "「표 확정」을 다시" in html


def test_채운_자리는_직접_입력으로_남는다():
    """자동이 낸 값과 사람이 넣은 값을 같은 얼굴로 두지 않는다."""
    html = _screen()
    # ★창을 «글자 수» 로 잡으면 안 된다 — 이 함수에 무리가 하나 더 붙는 순간
    #   찾던 것이 창 밖으로 밀려나 «없어졌다» 고 틀리게 실패한다(F-11d 에서
    #   실제로 그랬다). 함수의 끝을 경계로 잡는다.
    i = html.index("function collectIssues()")
    seg = html[i:html.index("function renderIssues()", i)]
    assert 'key: "applied"' in seg
    assert "직접 입력 — 사람이 채운 자리" in seg
    # 사유도 함께 보인다.
    assert "a.note" in seg


def test_등가길이는_쌍_단위로_채운다():
    """[§18 ②] 라이브러리 구멍은 (종류, 호칭경) 쌍이 단위다."""
    html = _screen()
    # ★창을 «글자 수» 로 잡으면 안 된다 — 이 함수에 무리가 하나 더 붙는 순간
    #   찾던 것이 창 밖으로 밀려나 «없어졌다» 고 틀리게 실패한다(F-11d 에서
    #   실제로 그랬다). 함수의 끝을 경계로 잡는다.
    i = html.index("function collectIssues()")
    seg = html[i:html.index("function renderIssues()", i)]
    assert 'type: "eq_len", kind: String(p.kind), dia: Number(p.dia)' in seg
    assert "한 쌍을 채우면" in seg


def test_미해결_목록이_표까지_실려_온다():
    """엔진이 남겨도 표가 안 들고 오면 화면은 여전히 개수만 본다."""
    from services.cad_import.design.tables import PipeTablesG

    t = PipeTablesG()
    assert hasattr(t, "unresolved")
    assert "unresolved" in t.as_dict()


def test_이상_목록은_잘라도_말한다():
    """조용히 자르지 않는다 — 몇 건을 안 보여주는지 적는다(저장소 규약)."""
    html = _screen()
    assert "const ISSUE_CAP" in html
    i = html.index("function renderIssues()")
    seg = html[i:i + 2600]      # 채우기 칸이 붙어 함수가 길어졌다
    assert "그 외" in seg and "rest" in seg


def test_이상_표시는_산출물을_안_건드린다():
    """표시 전용 증명 — 서버에 이 기능의 자리가 아예 없다."""
    import inspect

    from routes.module_f import api_design, views

    for mod in (api_design, views):
        src = inspect.getsource(mod)
        for banned in ("collectIssues", "dg_issues", "issues"):
            assert banned not in src, f"{mod.__name__} 에 {banned} 가 있다"


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
