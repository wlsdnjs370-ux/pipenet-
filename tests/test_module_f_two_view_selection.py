# -*- coding: utf-8 -*-
"""[두 화면 선정일치] 평면 보기와 수리계산 표가 같은 헤드를 그린다.

지시서 `ModuleF_두화면_선정일치_지시서.md`.

■ 증상 — 개수는 같은데 알맹이가 달랐다

  같은 도면·같은 세션에서 「평면에서 보기」를 켠 화면과 끈 화면의 헤드를
  좌표로 맞대 봤다(회전 없이 축척+평행이동만)::

      맞은 것        27 쌍       ①에만 4개       ②에만 4개
      헤드 총수      ① 30 · ② 30              ← **개수는 같다**

  기하는 무결했다. 다른 것은 «어느 헤드를 골랐는가» 하나뿐이었다.

■ 원인 — 선정이 두 벌이고, 뒤엣것을 아무도 화면에 돌려주지 않았다

      평면 보기   sess["worst"]   ← 손질이 고른 K개
      수리계산    got["worst"]    ← 못 붙는 것을 다음 순위로 채운 K개

  채우는 동작 자체는 옳다(기준개수 K 유지 · `a44ec64`). 잘못은 그 결과를
  **화면에 돌려주지 않고, 바뀌었다는 사실도 말하지 않은 것**이다. 그 침묵이
  세 곳에 있었다 — `got["_filled"]` 는 읽는 곳이 0곳이었고, 두 인계 함수는
  「다른 헤드로 채우지 않았습니다」라고 **거짓**을 적고 있었다.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _src(rel):
    return (_ROOT / rel).read_text(encoding="utf-8")


class _B:
    """판 흉내 — 이 시험이 보는 것은 `disks` 와 물길 상태뿐이다."""

    def __init__(self, disks, wet=None):
        self.disks = disks
        self._wet = wet

    def water_state(self):
        return {"wet_heads": list(self._wet if self._wet is not None
                                  else range(len(self.disks)))}


# ═══════════════════════════ §2-1 선정을 한 곳에서 정한다
def _got(heads, **kw):
    g = {"worst": {"heads": list(heads), "far_m": 1.0, "near_m": 0.5,
                   "edges": set(), "loads": {}, "reachable": len(heads)}}
    g.update(kw)
    return g


def test_표가_서면_세션_선정이_표와_같아진다():
    """★이 한 줄이 없어서 평면 화면이 계속 손질 선정을 그렸다."""
    from routes.module_f.api_design import _adopt_final_worst
    sess = {"worst": {"heads": [1, 2, 3], "zones": [[0, 0, 9, 9]],
                      "source_tag": "Z1", "sheet": 2, "candidates": 20},
            "worst_rev": ("옛 지문",)}
    out = _adopt_final_worst(sess, _got([1, 2, 9]))
    assert sess["worst"]["heads"] == [1, 2, 9], "표 선정을 안 받았다"
    assert out is sess["worst"]
    # 손질 원본은 사라지지 않는다 — «사람이 고른 것» 이 무엇이었는지 남는다.
    assert sess["worst_edit"]["heads"] == [1, 2, 3]


def test_손질만_아는_칸은_물려받는다():
    """★영역·급수원·장은 선정 계산이 모르는 값이다 — 안 물려주면 화면에서
    사각형이 사라지고 급수원 이름이 빈다(조용히 다르게 그린다)."""
    from routes.module_f.api_design import _adopt_final_worst
    sess = {"worst": {"heads": [1], "zones": [[0, 0, 9, 9]], "sheet": 2,
                      "source_tag": "Z2", "source_index": 1,
                      "candidates": 44}}
    _adopt_final_worst(sess, _got([7]))
    w = sess["worst"]
    assert w["zones"] == [[0, 0, 9, 9]] and w["sheet"] == 2
    assert w["source_tag"] == "Z2" and w["source_index"] == 1
    assert w["candidates"] == 44


def test_두_번째_확정이_손질_원본을_안_덮는다():
    """★멱등 — 두 번째 build 는 이미 «최종 선정» 을 받는다. 그때 원본을
    덮으면 「사람이 고른 것」이 영영 사라진다."""
    from routes.module_f.api_design import _adopt_final_worst
    sess = {"worst": {"heads": [1, 2, 3]}}
    _adopt_final_worst(sess, _got([1, 2, 9]))
    _adopt_final_worst(sess, _got([1, 2, 9]))
    assert sess["worst_edit"]["heads"] == [1, 2, 3]
    assert sess["worst"]["heads"] == [1, 2, 9]


def test_corridor_지문을_지운다():
    """★안 지우면 `_edit_state` 가 «안 바뀜» 으로 보고 옛 망을 그대로 둔다.

    1KB 규약(안 바뀐 블록은 안 싣는다)이 여기서는 «고친 것이 화면에 안
    보인다» 로 나타난다 — 서버는 맞췄는데 화면만 옛 그림을 든다.
    """
    from routes.module_f.api_design import _adopt_final_worst
    sess = {"worst": {"heads": [1]}, "worst_rev": ("옛 지문",)}
    _adopt_final_worst(sess, _got([2]))
    assert sess.get("worst_rev") is None


def test_표가_못_서면_선정을_안_바꾼다():
    from routes.module_f.api_design import _adopt_final_worst
    sess = {"worst": {"heads": [1, 2]}}
    assert _adopt_final_worst(sess, {"worst": {"heads": []}}) is None
    assert _adopt_final_worst(sess, {}) is None
    assert sess["worst"]["heads"] == [1, 2] and "worst_edit" not in sess


def test_손질에서_다시_고르면_그것이_최신이다():
    """사람이 다시 고른 것이 언제나 최신 — 접어 둔 원본은 버린다."""
    s = _src("routes/module_f/api_edit.py")
    i = s.index('sess["worst"] = w')
    assert 'sess["worst_edit"] = None' in s[i:i + 500], "원본을 안 버린다"


def test_선정이_지워지는_자리마다_원본도_지운다():
    """★한 자리라도 놓치면 다음 표가 옛 영역을 물려받는다."""
    for rel in ("routes/module_f/api_edit.py", "routes/module_f/api_pick.py"):
        s = _src(rel)
        n_clear = s.count('sess["worst"] = None')
        n_edit = s.count('sess["worst_edit"] = None')
        assert n_edit >= n_clear, f"{rel}: 선정만 지우고 원본을 남긴 자리가 있다"


def test_화면이_어느_선정을_보는지_말한다():
    """같은 그림에 두 뜻(손질이 고른 것 / 표가 쓴 것)이 있으면 판단을 못 한다.

    ★표시는 **선정 자신이** 든다 — 손질 원본의 유무로 미루면, 최불리를 안
      누르고 바로 확정한 세션에서 화면이 표 선정을 그리면서 아니라고 말한다.
    """
    from routes.module_f.api_design import _adopt_final_worst
    s = _src("routes/module_f/remote30.py")
    assert '"from_design": bool(w.get("from_design"))' in s
    js = _src("static/module_f.js")
    assert "w.from_design" in js, "화면이 그 값을 안 읽는다"
    # 손질을 안 거친 세션에서도 참이어야 한다(원본이 없다).
    sess: dict = {}
    _adopt_final_worst(sess, _got([3, 4]))
    assert sess["worst"]["from_design"] is True
    assert "worst_edit" not in sess
    # 두 번째 확정이 «표에서 온 선정» 을 원본이라 부르지 않는다.
    _adopt_final_worst(sess, _got([3, 5]))
    assert "worst_edit" not in sess


def test_표를_확정하면_손질_상태를_다시_받는다():
    """서버만 맞추고 화면이 안 받으면 고친 것이 하나도 안 보인다."""
    js = _src("static/module_f.js")
    i = js.index('const d = await post("/api/module-f/design/build"')
    seg = js[i:i + 1800]
    assert "/api/module-f/edit/state" in seg, "표 확정 뒤 손질 상태를 안 받는다"
    assert "setEdit(" in seg


# ═══════════════════════════ §2-2 바뀌었으면 말한다
def test_채웠으면_채웠다고_말한다():
    """★종전 문구는 「다른 헤드로 채우지 않았습니다」였다 — 채워 놓고."""
    from routes.module_f.api_design import _worst_handoff_note
    got = _got([1, 2, 9], candidate_heads=40)
    note = _worst_handoff_note(got, [1, 2, 3], 3, 3, filled=1,
                               board=_B([(0, 0, 4), (10, 0, 4), (20, 0, 4),
                                         (30, 0, 4)]))
    assert note["filled"] == 1
    assert any("다음 순위" in m for m in note["messages"]), note["messages"]
    assert not any("채우지 않았습니다" in m for m in note["messages"])
    assert [r["disk"] for r in note["swapped_out"]] == [3]
    assert [r["disk"] for r in note["swapped_in"]] == [9]


def test_안_채웠으면_문구가_그대로다():
    """★교체가 없는 세션은 응답도 문구도 **종전 그대로**여야 한다(기준 4).

    없는 일을 `filled: 0` 으로 적으면 그것만으로 응답이 달라진다 — 칸 자체를
    안 만든다.
    """
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note({"candidate_heads": 9}, list(range(12)), 12, 12)
    assert note["not_attachable"] == 3
    assert any("다른 헤드로 채우지 않았습니다" in m for m in note["messages"])
    for k in ("filled", "swapped_out", "swapped_in"):
        assert k not in note, k


def test_다_그대로_받으면_새_칸이_안_생긴다():
    """★교체가 없는 «정상» 세션 — 응답에 새 이름이 하나도 안 붙는다."""
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note(_got([1, 2, 3], candidate_heads=3),
                               [1, 2, 3], 3, 3,
                               board=_B([(0, 0, 4)] * 4))
    assert set(note) == {"picked", "k", "candidates", "from_edit",
                         "not_attachable", "messages"}
    assert not note["messages"]


def test_바뀐_헤드는_두_집합의_차로_낸다():
    """★새로 세지 않는다 — 이미 있는 두 집합의 차다."""
    from routes.module_f.api_design import _worst_handoff_note
    board = _B([(i * 10.0, 0.0, 4.0) for i in range(10)])
    note = _worst_handoff_note(_got([0, 5, 6], candidate_heads=9),
                               [0, 1, 2], 3, 3, filled=2, board=board,
                               reasons={1: "center_dry", 2: "chord_only"})
    assert [r["disk"] for r in note["swapped_out"]] == [1, 2]
    assert [r["disk"] for r in note["swapped_in"]] == [5, 6]
    # 자리와 사유가 함께 나온다 — 번호만 주면 도면에서 못 찾는다.
    assert note["swapped_out"][0]["xy"] == [10.0, 0.0]
    assert note["swapped_out"][0]["why"] == "center_dry"
    assert "물길" in note["swapped_out"][0]["why_text"]


def test_개수가_같아도_교체가_있으면_표를_다시_본다():
    """★종전에는 `lost <= 0` 하나로 막혀 여기서 통째로 조용해졌다."""
    from routes.module_f.api_design import _handoff_after_table

    class _T:
        nodes = [{"label": "1", "x": 1000.0, "y": 1000.0}]
        nozzles = [{"in": "1"}]

    got = {"handoff": {"from_edit": True, "picked": 1, "filled": 1,
                       "messages": []},
           "_picked": [1], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, _T(), _B([(1000.0, 1000.0, 40.0),
                                        (9000.0, 9000.0, 40.0)]))
    h = got["handoff"]
    assert h["missing"] == 0                     # 개수는 K 그대로다
    assert h["missing_by_table"] == 1, h          # 그런데 알맹이가 바뀌었다
    # 거짓 문구를 새로 달지 않는다 — 그 말은 §2-2 가 이미 했다.
    assert not any("채우지 않았습니다" in m for m in h["messages"]), h


def test_교체도_결손도_없으면_아무_말도_안_한다():
    from routes.module_f.api_design import _handoff_after_table

    class _T:
        nodes = [{"label": "1", "x": 1000.0, "y": 1000.0}]
        nozzles = [{"in": "1"}]

    got = {"handoff": {"from_edit": True, "picked": 1, "messages": []},
           "_picked": [0], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, _T(), _B([(1000.0, 1000.0, 40.0)]))
    assert got["handoff"]["missing"] == 0
    assert not got["handoff"]["messages"]
    assert "missing_by_table" not in got["handoff"]


def test_화면이_바뀐_헤드를_말한다():
    js = _src("static/module_f.js")
    i = js.index("function handoffLines(")
    seg = js[i:i + 2600]
    assert "바뀐 헤드" in seg and "why_text" in seg
    assert "다음 순위로" in seg, "채웠다는 사실을 화면이 안 말한다"


# ═══════════════════════════ §2-3 빠진 헤드를 도면에 찍는다
def test_빠진_헤드를_따로_켤_수_있다():
    """★`unattached` 는 도면 전체가 대상이라 수백 개가 될 수 있다 — 내가 고른
    것 중 빠진 넷이 그 사이에 묻히면 켜도 못 찾는다."""
    from routes.module_f.api_design import _classify_excluded
    got = {"handoff": {"swapped_out": [{"disk": 2, "xy": [7.0, 8.0],
                                        "why": "center_dry"}]}}
    board = _B([(0, 0, 4), (10, 0, 4), (7, 8, 4)])
    out = _classify_excluded({"edit": None}, got, board,
                             probe={"ok": True, "wet": {0, 1},
                                    "reason": {2: "center_dry"}})
    assert out["swapped_out"]["n"] == 1
    assert out["swapped_out"]["xy"] == [[7.0, 8.0]]
    # 「이음 끊김」 안에서도 갈래를 센다.
    assert out["unattached"]["n"] == 1
    assert out["unattached"]["why"] == {"center_dry": 1}


def test_탐침을_다시_안_잰다():
    """★종전에는 이 함수가 제 손으로 전체망 전개를 한 번 더 돌렸다(117초).

    게다가 `selected_source` 를 안 넘겨 **다른 급수원 기준**의 답이 나올 수
    있었다 — 화면의 사유가 표와 다른 말을 하는 자리다. `sess["edit"]` 가
    None 인데도 도는 것이 그 증거다(다시 잰다면 여기서 죽는다).
    """
    from routes.module_f.api_design import _classify_excluded
    out = _classify_excluded({"edit": None}, {}, _B([(0, 0, 4)]),
                             probe={"ok": True, "wet": {0}, "reason": {}})
    assert out["total"] == 1 and out["unattached"]["n"] == 0
    s = _src("routes/module_f/api_design.py")
    assert "_classify_excluded(sess, got, es.board, probe=probe)" in s


def test_화면에_체크박스와_색이_있다():
    html = _src("templates/module_f.html")
    assert 'id="dg-mk-swap"' in html and 'id="dg-swap-why"' in html
    js = _src("static/module_f.js")
    assert '["swapped_out", "dg-mk-swap"' in js
    assert '"dg-mk-swap"' in js.split("function designMarksOn")[1][:200]


# ═══════════════════════════ §2-4 왜 못 붙었는지 한 줄로 가른다
def test_여섯_갈래가_한_곳에_있다():
    from services.cad_import.convert.planar import HEAD_REASON_TEXT
    assert set(HEAD_REASON_TEXT) == {
        "dry", "no_center", "center_dry", "chord_only", "pass_under", "shared"}
    for v in HEAD_REASON_TEXT.values():
        assert len(v) > 10, v          # «한 줄» 은 사람이 읽는 문장이다


def test_전개가_사유를_내보낸다():
    s = _src("cad_project_editor_g/services/cad_import/convert/planar.py")
    assert '"head_reason": head_reason' in s
    r = _src("cad_project_editor_g/services/cad_import/design/restrict.py")
    assert '"reason": dict(built.get("head_reason") or {})' in r


def test_접속_판정이_사유를_적는다():
    """★규칙은 한 글자도 안 바꾼다 — 이미 갈린 자리에서 이름표만 딴다.

    기하로 직접 확인한다(원 반지름 100 · 붙었다 자 50mm):
      ① 아무것도 없는 헤드                → no_center
      ② 양끝이 다 원 위인 «현»(문양)       → chord_only
      ③ 테두리를 «지나가는» 관(끝이 아님)  → pass_under
      ④ 중심에 노드가 있으면 사유가 없다   → 붙었다
    """
    from services.cad_import.pipeline.flow import attach_heads_center

    # ① 헤드만 있고 근처에 아무 노드도 없다
    pts = [(50000.0, 0.0), (51000.0, 0.0)]
    edges = {(0, 1)}
    why: dict = {}
    _p, _e, ctr, _n, _m = attach_heads_center(pts, edges, [(0.0, 0.0, 100.0)],
                                              why=why)
    assert ctr == [None] and why == {0: "no_center"}

    # ② 양끝이 다 원 위 — 하향식 기호의 가로막대다(팔이 아니다)
    pts = [(100.0, 0.0), (0.0, 100.0)]
    edges = {(0, 1)}
    why = {}
    _p, _e, ctr, _n, _m = attach_heads_center(pts, edges, [(0.0, 0.0, 100.0)],
                                              why=why)
    assert ctr == [None] and why == {0: "chord_only"}

    # ③ 관이 테두리를 지나간다 — 끝나는 자리가 아니다
    pts = [(-300.0, 0.0), (100.0, 0.0), (300.0, 0.0)]
    edges = {(0, 1), (1, 2)}
    why = {}
    _p, _e, ctr, _n, _m = attach_heads_center(pts, edges, [(0.0, 0.0, 100.0)],
                                              why=why)
    assert ctr == [None] and why == {0: "pass_under"}

    # ④ 중심에 간선 달린 노드가 있으면 그것이 접속점이다
    pts = [(0.0, 0.0), (0.0, 400.0)]
    edges = {(0, 1)}
    why = {}
    _p, _e, ctr, _n, _m = attach_heads_center(pts, edges, [(0.0, 0.0, 100.0)],
                                              why=why)
    assert ctr == [0] and why == {}


def test_사유를_안_물으면_종전과_같다():
    """★`why` 를 안 주면(기본) 동작이 한 글자도 안 달라진다."""
    from services.cad_import.pipeline.flow import attach_heads_center
    args = ([(100.0, 0.0), (0.0, 100.0)], {(0, 1)}, [(0.0, 0.0, 100.0)])
    a = attach_heads_center(*args)
    b = attach_heads_center(*args, why={})
    assert a[2] == b[2] and a[3] == b[3] and a[4] == b[4]
    assert set(a[1]) == set(b[1]) and list(a[0]) == list(b[0])


# ═══════════════════════════ §2-5 «채우지 않기» 를 고를 수 있다
def test_기본은_채움이다():
    """★기본을 바꾸면 옛 세션의 산출이 조용히 달라진다 — 스위치만 둔다."""
    from routes.module_f.api_design import _DEFAULT_SETTINGS
    assert _DEFAULT_SETTINGS["fill_short"] is True


def test_옛_세션에도_새_칸이_채워진다():
    """★없으면 `cfg["fill_short"]` 가 KeyError 로 죽는다 — 세션은 오래 산다."""
    from routes.module_f.api_design import _settings
    sess = {"design_settings": {"k": 30, "schedule": "KSD 3507",
                                "iso": True, "iso_z_scale": 1.0,
                                "canvas_units": 3000.0, "lift_ref": "valve",
                                "head_stub_pct": 2.5}}
    cur = _settings(sess, {})
    assert cur["fill_short"] is True
    assert _settings({}, {"fill_short": False})["fill_short"] is False


def test_끄면_막고_말한다():
    """§2-5 — 「사람이 고른 것만 쓴다」를 고르면 K 미달을 그대로 보고한다."""
    s = _src("routes/module_f/api_design.py")
    i = s.index("if wet and short < k_use:")
    seg = s[i:i + 700]
    assert 'if not cfg.get("fill_short", True):' in seg
    assert '"ok": False' in seg and "멈춥니다" in seg
    html = _src("templates/module_f.html")
    assert 'id="dg-fill" checked' in html, "기본이 «켬» 이 아니다"
    assert "fill_short: $(\"dg-fill\").checked" in _src("static/module_f.js")


# ═══════════════════════════ 영역 넘어감 — 채움은 한 구역 안에서
#
#   2026-09-11 사용자 지적: 「대명동 기준, 영역1 내에 있는 2개 헤드가 영역2 및
#   배관망을 강제로 넘어간다.」 실측으로 재현했다(`data/_probe_zone_cross.py`
#   · 영역1 을 K=10 에 딱 맞게 좁힌 판):
#
#       손질 선정 10개 전부 영역1 → 2개가 안 붙음(pass_under)
#       채움  헤드 47 → 영역1 · 헤드 82 → **영역2**(x 284,092 · 반대편)
#       표 최종 영역1 9 · 영역2 1 · corridor 두 영역에 걸침 · 배관 78 → 120
#
#   설계면적은 «하나의 방호구역 안에서 인접한 K개» 다. 사람이 사각형을 둘로
#   나눠 그린 것은 「이 둘은 다른 구역」이라는 뜻이지 합쳐 달라는 뜻이 아니다.
_Z1 = [0.0, 0.0, 100.0, 100.0]
_Z2 = [900.0, 0.0, 1000.0, 100.0]
_DISKS = [(10.0, 10.0, 4.0),     # 0 영역1
          (20.0, 20.0, 4.0),     # 1 영역1
          (30.0, 30.0, 4.0),     # 2 영역1
          (950.0, 50.0, 4.0),    # 3 영역2
          (960.0, 60.0, 4.0),    # 4 영역2
          (500.0, 500.0, 4.0)]   # 5 어느 영역도 아님


def test_영역을_안_그렸으면_종전_그대로():
    """★영역 없는 세션은 한 줄도 안 달라져야 한다(골든이 그 경로다)."""
    from routes.module_f.api_design import zone_confined_pool
    pool = [0, 1, 3, 4]
    assert zone_confined_pool(pool, [0], None, _DISKS) is pool
    assert zone_confined_pool(pool, [0], [], _DISKS) is pool


def test_선정이_든_영역_밖은_버린다():
    """★이것이 사용자가 본 그 버그다 — 영역1 자리를 영역2 가 채웠다."""
    from routes.module_f.api_design import zone_confined_pool
    got = zone_confined_pool([0, 1, 2, 3, 4], [0, 1], [_Z1, _Z2], _DISKS)
    assert got == [0, 1, 2], got          # 영역2(3·4)는 후보에서 빠진다


def test_두_영역에_걸쳐_뽑았으면_둘_다_남긴다():
    """사람이 그렇게 고른 것이다 — 그때는 합집합이 맞다."""
    from routes.module_f.api_design import zone_confined_pool
    got = zone_confined_pool([0, 1, 3, 4], [0, 3], [_Z1, _Z2], _DISKS)
    assert got == [0, 1, 3, 4]


def test_가둘_근거가_없으면_손대지_않는다():
    """선정이 어느 사각형에도 안 들어가면(있을 수 없지만) 조용히 바꾸지 않는다."""
    from routes.module_f.api_design import zone_confined_pool
    pool = [0, 1, 3]
    assert zone_confined_pool(pool, [5], [_Z1, _Z2], _DISKS) is pool


def test_채울_때_영역을_가두고_모자라면_말한다():
    s = _src("routes/module_f/api_design.py")
    i = s.index("if wet and short < k_use:")
    seg = s[i:i + 1600]
    assert "zone_confined_pool(" in seg, "채움이 영역을 안 가둔다"
    assert "avail < k_use" in seg and "다른 영역에서 끌어오지 않습니다" in seg


def test_채웠는데_모자라면_채우지_않았다고_안_한다():
    """★채워 놓고 「채우지 않았습니다」라고 적으면 거짓이다."""
    from routes.module_f.api_design import _handoff_after_table

    class _T:
        nodes = [{"label": "1", "x": 1000.0, "y": 1000.0}]
        nozzles = [{"in": "1"}]

    got = {"handoff": {"from_edit": True, "picked": 2, "filled": 1,
                       "messages": []},
           "_picked": [0, 1], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, _T(), _B([(1000.0, 1000.0, 40.0),
                                        (9000.0, 9000.0, 40.0)]))
    msgs = got["handoff"]["messages"]
    assert msgs and "채웠지만" in msgs[0], msgs
    assert not any("채우지 않았습니다" in m for m in msgs), msgs
    js = _src("static/module_f.js")
    assert "그 영역 안에 더는 없습니다" in js, "화면이 그 갈래를 안 가른다"


def test_표가_서기_전에_기준개수를_단정하지_않는다():
    """★제한 전개에서 또 떨어질 수 있다(실측 선정 10 → 표 9) — 단정해 두면
    뒤이어 붙는 «9개 왔습니다» 와 서로 어긋난다."""
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note(_got([1, 2, 9], candidate_heads=40),
                               [1, 2, 3], 3, 3, filled=1,
                               board=_B([(0, 0, 4)] * 10))
    assert not any("지켰습니다" in m for m in note["messages"]), note["messages"]
    assert any("선정은 3개를 채웠습니다" in m for m in note["messages"])


# ═══════════════════════════ 표가 «옛 것» 이면 말한다
#
#   2026-09-11 사용자 지적: 「최불리 배관망에서 등록했던 배관망을 그대로
#   아이소매트릭에 가져오는 것조차 안 된다.」 실측으로 재현했다
#   (`data/_probe_stale_iso.py`):
#
#       최불리 A(왼쪽 영역 K=8) → 표 확정 → 노즐 8
#       최불리 B(오른쪽 영역 K=8) → **표를 안 누름**
#       평면 보기 = 새 선정 B · 아이소 = 옛 표 A · 겹치는 헤드 **0개**
#       화면은 아무 말도 안 했다.
#
#   표는 «표 확정을 누른 그 순간» 의 사진이다. 그 뒤에 선정이나 손질판이
#   바뀌면 사진은 옛 것이 된다 — 그 사실을 화면이 말해야 한다.
class _ES:
    def __init__(self, board):
        self.board = board


def _sess_with(heads, zones=None, disks=None, joins=0):
    b = _B(disks if disks is not None else [(0, 0, 4)] * 5)
    b.joins = [0] * joins
    b.deletes = []
    b.edges = [(0, 1)]
    b.pts = [(0.0, 0.0), (1.0, 1.0)]
    b.sources = [0]
    b.valves = [0]
    b.disk_kinds = ["하향식"] * len(b.disks)
    return {"worst": {"heads": list(heads), "zones": zones or [],
                      "source_tag": "Z1", "sheet": None},
            "edit": _ES(b)}


def test_선정이_바뀌면_지문이_바뀐다():
    from routes.module_f.api_design import _selection_sig
    a = _selection_sig(_sess_with([1, 2, 3]))
    assert _selection_sig(_sess_with([1, 2, 3])) == a      # 같으면 같다
    assert _selection_sig(_sess_with([1, 2, 9])) != a      # 헤드
    assert _selection_sig(_sess_with([1, 2, 3], zones=[[0, 0, 9, 9]])) != a
    assert _selection_sig(_sess_with([1, 2, 3], joins=1)) != a   # 손질판


def test_표가_옛것이면_무엇이_달라졌는지_말한다():
    from routes.module_f.api_design import _selection_sig, _design_stale
    sess = _sess_with([1, 2, 3])
    sess["design"] = {"sig": _selection_sig(sess)}
    assert _design_stale(sess) is None                     # 갓 만든 표
    sess["worst"]["heads"] = [4, 5, 6]                     # 다시 골랐다
    st = _design_stale(sess)
    assert st and st["heads_in_table"] == 3 and st["heads_now"] == 3
    assert any("최불리 선정이 바뀌었습니다" in w for w in st["why"]), st


def test_손질을_고쳐도_옛것이_된다():
    from routes.module_f.api_design import _selection_sig, _design_stale
    sess = _sess_with([1, 2, 3])
    sess["design"] = {"sig": _selection_sig(sess)}
    sess["edit"].board.joins = [0]                         # 이음 하나 추가
    st = _design_stale(sess)
    assert st and any("손질에서 배관망을 고쳤습니다" in w for w in st["why"]), st


def test_표가_없으면_옛것도_아니다():
    from routes.module_f.api_design import _design_stale
    assert _design_stale(_sess_with([1])) is None          # 아직 표가 없다
    sess = _sess_with([1])
    sess["design"] = {"tables": None}                      # 옛 세션(지문 없음)
    assert _design_stale(sess) is None


def test_표를_만들_때_지문을_박는다():
    s = _src("routes/module_f/api_design.py")
    i = s.index('sess["design"] = {')
    assert '"sig": _selection_sig(sess)' in s[i:i + 400], "지문을 안 박는다"
    assert '"stale": _design_stale(sess)' in s, "preview 가 안 알린다"


def test_화면이_옛것이라고_말한다():
    js = _src("static/module_f.js")
    assert "function renderStale(" in js
    assert "renderStale(d.stale)" in js, "preview 응답을 안 읽는다"
    assert "지금 보이는 표·아이소는 옛 것입니다" in js
    assert "「표 확정」을 다시 눌러야" in js, "무엇을 하면 되는지 안 말한다"
    assert 'id="dg-stale"' in _src("templates/module_f.html")


# ═══════════════════════════ 배관망도 한 벌로 — 평면이 «표의 망» 을 그린다
#
#   2026-09-11 사용자 지적(대명동 · K=30 · 영역 2곳): 선정은 맞췄는데
#   **배관망이 달랐다.** 실측(`scripts/_probe_corridor_vs_iso.py --k 30 --zones`):
#
#       corridor 461선분 156.1 m  /  표 202선분 160.6 m
#       표에만 있는 긴 배관 3.35·3.00·2.96·2.34·2.03 m … (합 11.1 m)
#
#   ★막다른 관이 남아서가 아니다 — 표의 막다른 끝은 5곳(0.03~0.50 m)뿐이고
#     2곳은 급수원 스텁, 3곳은 「끝배관 1단 보호」다(전개 로그와 일치).
#     원인은 **두 그림이 서로 다른 그래프에서 최단경로를 뽑는 것**이었다.
#   조치 뒤: 평면 202선분 160.6 m · 양쪽 차이 0.
def test_표가_있으면_평면이_표의_망을_그린다():
    s = _src("routes/module_f/remote30.py")
    assert "def _design_corridor(" in s
    i = s.index('"corridor": (dc["corridor"] if dc else')
    assert i > 0, "평면 corridor 가 표를 안 본다"
    assert '"net_from": ("design" if dc else "edit")' in s


def test_옛_표로는_안_그린다():
    """★옛 표를 «지금 망» 으로 그리면 그것이 바로 고치려던 증상이다."""
    s = _src("routes/module_f/remote30.py")
    i = s.index("def _design_corridor(")
    seg = s[i:i + 2600]
    assert "if _design_stale(sess):" in seg and "return None" in seg


def test_최원_경로도_표_위에서_잇는다():
    """그린 망과 다른 줄을 덧그리면 더 헷갈린다."""
    s = _src("routes/module_f/remote30.py")
    assert '"worst_path": (dc["path"] if dc and dc["path"] else' in s


def test_화면이_어느_망인지_말한다():
    js = _src("static/module_f.js")
    assert 'w.net_from === "design"' in js
    assert "표와 같은 배관망" in js


# ═══════════════════════════ §5 금지 사항 — 되돌아가지 않는다
def test_백필을_안_없앴다():
    """★없애면 기준개수 K 가 다시 무너진다(`a44ec64` 가 고친 그 증상)."""
    s = _src("routes/module_f/api_design.py")
    i = s.index("filled = 0")
    seg = s[i:i + 1200]
    assert "if wet and short < k_use:" in seg
    assert "only = (set(pool) & wet) if pool else None" in seg


def test_접속_판정_규칙은_그대로다():
    """§2-4 는 «가르기» 지 «고치기» 가 아니다 — 규칙을 손대면 도면 전체의
    물닿음 수가 움직이고 그 영향은 이 지시서 밖이다."""
    s = _src("cad_project_editor_g/services/cad_import/pipeline/flow.py")
    assert "elif abs(d - hr) <= tol and len(nbr[n]) == 1:" in s
    assert "if abs(do - hr) <= 2.0:" in s
    assert "if d <= ARM_CTR or abs(d - hr) <= tol:" in s      # head_nodes
