# -*- coding: utf-8 -*-
"""[신축배관 접기] 한 가닥을 직선 하나로 — `ModuleF_신축배관_접기_지시서.md`.

■ 어제의 「빼기」가 왜 철회됐나

  빼면 대명동에서 물닿음 헤드가 **111 → 5** 다(실측). 도면의 후렉시블은
  가지관 → 헤드를 **평면에서 수평으로** 잇는 300~1,000 mm 구간이고, 변환이
  만드는 접속관은 헤드를 **수직으로** 세우는 0.3 m 다 — 중복이 아니다.
  그래서 `exclude-layer` 는 아예 없앴다: 남겨 두면 언젠가 누가 누른다.

■ ★한 가닥은 «선분 + 호» 다 — 이 파일이 지키는 첫째 사실

  대명동 실측 (`scripts/_probe_flex_world.py`)::

      선분  320개 · 44,554 mm      ← `Board.by_bundle` 에 들어가는 것
      호    115개 · 14,057 mm      ← **아무 데도 안 들어간다**

      선분만으로 사슬을 이으면   200토막 (4~375 mm · 중앙값 200)
      호를 현으로 넣어 이으면     85가닥 (663~723 mm · 중앙값 665 · 합 57,271)

  85 · 663~723 · 665 · 57,271 은 지시서가 원도면을 직접 재서 낸 값과 같다.
  즉 「가닥당 4~9 선분」은 선분과 호를 합한 조각 수였다.

■ ★가닥의 끝은 헤드만이 아니다 — 둘째 사실

      151(차수1·헤드) – 152 – 153 – 128(차수3·가지관) – 154   ← 꼬리 30 mm

  128 을 삼키면 거기 붙은 가지관이 통째로 떨어진다. 그래서 «다른 배관이 붙은
  노드» 에서 **자르고** 구간마다 접는다. 처음에 가닥 전체를 접으려 했다가
  85가닥 중 79가 거절됐다.

■ ★길이는 좌표가 아니라 «선언» 이 권위다 — 셋째 사실

      L1(맨해튼) 합 65,990 → 현의 L1 60,017   (−5,974 mm · −9.1 %)

  85가닥 중 38이 되돌아 꺾여 그렇다. 줄면 마찰손실 과소 = **비보수측**이다.

■ ★선언은 «어디서 · 어떤 자로» 얹는가 — 넷째·다섯째 사실

  ⑴ 배관을 만들 때 얹으면 안 된다. 바로 뒤 «마무리 1/2 — 배관 길이 갱신»
     (`update_all_pipe_lengths`)이 좌표에서 전부 다시 계산해 덮는다 — 실측으로
     선언이 표에 **한 건도** 안 실렸다. 지시서 §2 가 「좌표에서 길이를 다시
     계산하는 자리가 하나라도 남으면 이 작업의 의미가 없다」고 한 그 자리다.

  ⑵ 그 자로 재야 한다. 갱신이 쓰는 `compute_length` 는 `dist_3d` — **유클리드**
     다. `planar` 가 배관 생성 때 쓰는 L1 은 곧바로 덮여 효과가 없다. 한 번
     L1 으로 선언했다가 되돌렸다(그때 견준 「표 8,819 vs 좌표 L1 11,211」은
     자가 서로 달라 성립하지 않는 비교였다).

  ⑶ 넣을 때는 `graph.update_pipe()` 로. 속성에 직접 대입하면 색인·사본이 안
     따라와 `to_dict()` 가 옛 값을 낸다.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _ff():
    from routes.module_f import flexfold
    return flexfold


class _W:
    """세계 시늉 — 선분과 호를 함께 든다."""

    def __init__(self, segs=(), arcs=(), arc_ang=()):
        self.segs = list(segs)
        self.arcs = list(arcs)
        self.arc_ang = list(arc_ang)


class _B:
    def __init__(self, pts, edges):
        self.pts = list(pts)
        self.edges = frozenset(tuple(sorted(e)) for e in edges)


# ─────────────────────────────── ① 조각 — 호를 빠뜨리지 않는다
def test_호도_조각으로_센다():
    """★`Board.by_bundle` 은 선분만 담는다 — 호를 안 보면 가닥을 못 만든다."""
    ff = _ff()
    w = _W(segs=[("FX", 7, (0.0, 0.0), (10.0, 0.0))],
           arcs=[("FX", 7, 10.0, 10.0, 10.0)],
           arc_ang=[(-90.0, 90.0)])
    got = ff.flex_pieces(w, ["FX"])
    assert len(got) == 2, got
    # 호는 «현» 으로 들어간다 — (10,0) → (20,10)
    (a, b) = got[1]
    assert abs(a[0] - 10.0) < 1e-6 and abs(a[1] - 0.0) < 1e-6, a
    assert abs(b[0] - 20.0) < 1e-6 and abs(b[1] - 10.0) < 1e-6, b


def test_다른_레이어는_안_가져온다():
    ff = _ff()
    w = _W(segs=[("FX", 7, (0.0, 0.0), (10.0, 0.0)),
                 ("가지관", 6, (0.0, 5.0), (10.0, 5.0))])
    assert len(ff.flex_pieces(w, ["FX"])) == 1


# ─────────────────────────────── ② 가닥 — 격자가 아니라 eps 로 뭉친다
def test_선분과_호가_한_가닥으로_이어진다():
    ff = _ff()
    pieces = [((0.0, 0.0), (10.0, 0.0)),
              ((10.0, 0.0), (20.0, 10.0)),      # 호의 현
              ((20.0, 10.0), (20.0, 20.0))]
    chains, tangled, rest = ff.strands(pieces)
    assert len(chains) == 1 and not tangled and not rest
    assert len(chains[0]) == 4


def test_칸_경계에_걸친_이음도_붙는다():
    """★격자 반올림은 «칸 경계» 에서 갈라진다 — 한 폴리라인이 토막 난다.

    아래 두 조각은 5.0 을 사이에 두고 1e-7 만큼 어긋나 있다. `int(x // 0.5)`
    로만 칸을 나누면 9 와 10 으로 갈려 **다른 점**이 되고, 그러면 한 가닥이
    둘로 쪼개진다(실측: 85가닥이 196토막). 이웃 칸까지 보는 eps 클러스터는
    붙인다.
    """
    ff = _ff()
    pieces = [((0.0, 0.0), (4.9999999, 0.0)), ((5.0, 0.0), (9.0, 0.0))]
    chains, _t, rest = ff.strands(pieces)
    assert len(chains) == 1 and len(chains[0]) == 3, (chains, rest)


def test_갈래는_가닥이_아니다():
    """차수 3 이 섞이면 사슬이 아니다 — 지어내지 않고 «사슬 아님» 으로 센다."""
    ff = _ff()
    pieces = [((0.0, 0.0), (10.0, 0.0)),
              ((10.0, 0.0), (20.0, 0.0)),
              ((10.0, 0.0), (10.0, 10.0))]
    chains, tangled, _r = ff.strands(pieces)
    assert tangled, (chains, tangled)


# ─────────────────────────────── ③ 접기 — 분기에서 자른다
def test_한_가닥이_직선_하나가_된다():
    ff = _ff()
    b = _B([(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)],
           [(0, 1), (1, 2)])
    chain = [(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)]
    got = ff.fold_board(b, [chain])
    assert got["folded"] == 1 and got["skipped"] == 0, got
    assert b.edges == frozenset({(0, 2)}), b.edges


def test_접어도_길이는_접기_전_총연장이다():
    """★현(14.14)이 아니라 총연장(20)을 선언한다 — 줄이면 비보수측이다."""
    ff = _ff()
    b = _B([(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)], [(0, 1), (1, 2)])
    got = ff.fold_board(b, [[(0.0, 0.0), (10.0, 0.0), (10.0, 10.0)]])
    dec = got["edge_len_mm"][(0, 2)]
    assert abs(dec - 20.0) < 1e-6, dec
    assert dec > math.dist((0.0, 0.0), (10.0, 10.0))


def test_선언은_표와_같은_자로_잰다_유클리드():
    """★표의 길이 규약은 **유클리드** 다 — 다른 자로 재면 접힌 것만 어긋난다.

    `planar` 가 배관을 만들 때 L1 으로 재는 것은 맞지만, 그 값은 바로 뒤
    «마무리 1/2 — 배관 길이 갱신»(`compute_length` → `dist_3d`)이 유클리드로
    전부 덮는다. 한 번 L1 으로 선언했다가 되돌린 자리다.

    아래 가닥은 45° 대각 두 도막이다 — 유클리드 합 28.28 (L1 이면 40).
    """
    ff = _ff()
    path = [(0.0, 0.0), (10.0, 10.0), (20.0, 20.0)]
    assert abs(ff.run_length(path) - 20.0 * math.sqrt(2)) < 1e-6,         ff.run_length(path)


def test_다른_배관이_붙은_노드에서_자른다():
    """★그 노드를 삼키면 거기 붙은 가지관이 통째로 떨어진다(실측 노드 128)."""
    ff = _ff()
    #  0 – 1 – 2 – 3 이 가닥이고, 2 에 가지관 4 가 붙어 있다.
    b = _B([(0.0, 0.0), (10.0, 0.0), (20.0, 0.0), (20.0, 10.0),
            (20.0, -10.0)],
           [(0, 1), (1, 2), (2, 3), (2, 4)])
    got = ff.fold_board(b, [[(0.0, 0.0), (10.0, 0.0), (20.0, 0.0),
                             (20.0, 10.0)]])
    assert got["cuts"] == 1, got
    # 가지관 2–4 는 살아 있어야 한다.
    assert (2, 4) in b.edges, b.edges
    # 0–2 는 접히고, 2–3 은 이미 직선 하나라 그대로다.
    assert (0, 2) in b.edges and (2, 3) in b.edges, b.edges
    assert abs(got["edge_len_mm"][(0, 2)] - 20.0) < 1e-6


def test_이미_직선이면_안_건드린다():
    ff = _ff()
    b = _B([(0.0, 0.0), (10.0, 0.0)], [(0, 1)])
    got = ff.fold_board(b, [[(0.0, 0.0), (10.0, 0.0)]])
    assert got["folded"] == 0 and b.edges == frozenset({(0, 1)})


def test_못_찾은_점은_세어서_돌려준다():
    """지어내지 않는다 — 왜 못 접었는지가 남아야 산출 기록에 적을 수 있다."""
    ff = _ff()
    b = _B([(0.0, 0.0), (10.0, 0.0)], [(0, 1)])
    got = ff.fold_board(b, [[(0.0, 0.0), (10.0, 0.0), (9999.0, 9999.0)]])
    assert got["skipped"] == 1 and got["why"], got


# ─────────────────────────────── ④ 선언 길이가 좌표를 이긴다
def test_planar_가_선언_길이를_쓴다():
    """★좌표에서 다시 재는 자리가 하나라도 남으면 접기가 길이를 잃는다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "convert" / "planar.py").read_text(encoding="utf-8")
    i = src.index("if declared:")
    seg = src[i:i + 1200]
    # ★그래프의 정식 갱신 경로로 넣어야 `to_dict()` 가 새 값을 낸다.
    assert "editor.graph.update_pipe(_pid, length_m=" in seg, seg[:400]
    assert "declared_pipes.append" in seg, "선언을 쓴 배관을 안 남긴다"
    # 새 통로를 파지 않았다 — `_snap_origin`/`edge_ref` 가 나르던 그 키를 쓴다.
    assert "edge_len_mm" in src and "_snap_origin" in src


def test_선언은_덮는_자리_셋_모두_뒤에_얹는다():
    """★좌표에서 길이를 다시 계산하는 자리가 뒤에 하나라도 남으면 조용히 덮인다.

    실측으로 두 번 걸렸다 — 선언이 표에 **한 건도** 안 실렸다:

        ① 배관 생성 때 얹기  → «마무리 1/2 — 배관 길이 갱신» 이 덮음
        ② 마무리 뒤에 얹기   → `validate_and_fix_integrity(strict=True)` 가
                              내부에서 그 갱신을 **또** 부름

    지시서 §2 가 「좌표에서 길이를 다시 계산하는 자리가 하나라도 남으면 이
    작업의 의미가 없다」고 한 그 자리다. 그래서 `to_dict()` 바로 앞이다.
    """
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "convert" / "planar.py").read_text(encoding="utf-8")
    i_apply = src.index("if declared:")
    for name in ("update_all_pipe_lengths_in_pipe_data",
                 "validate_and_fix_integrity",
                 "edge_ref.pop(_pid, None)"):
        assert i_apply > src.index(name), f"선언을 «{name}» 앞에서 얹고 있다"
    # 그리고 그 뒤에 길이를 다시 재는 자리가 없어야 한다.
    tail = src[i_apply:]
    i_dict = tail.index("kfp = editor.to_dict()")
    assert "update_all_pipe_lengths" not in tail[:i_dict], tail[:i_dict][-300:]
    assert "validate_and_fix_integrity" not in tail[:i_dict]


def test_병합된_배관은_안_건드린다():
    """★노드정리가 접힌 배관을 옆과 병합했으면 선언만 얹는 순간 옆 몫을 잃는다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "convert" / "planar.py").read_text(encoding="utf-8")
    i = src.index("if declared:")
    seg = src[i:i + 1200]
    assert "_merged" in seg and "건너뜀" in seg, seg[:400]


def test_선언이_없으면_종전_그대로다():
    """§5 기준 6 — 접지 않은 상태의 산출은 한 바이트도 안 바뀐다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "convert" / "planar.py").read_text(encoding="utf-8")
    # 길이 계산은 종전 그대로 좌표 L1 이다 — 선언은 그 «뒤에» 덮을 뿐이다.
    assert "pipe.length_m = round(sum(abs(na[k] - nb[k]) for k in range(3)), 3)"         in src, "기본 길이 계산이 바뀌었다"
    # 선언이 비면 `if declared:` 블록을 통째로 건너뛴다.
    j = src.index("declared = {}")
    assert "edge_len_mm or {}" in src[j:j + 200]


def test_선언_길이가_payload_로_흐른다():
    """찍기 → 손질 → 전개까지 한 줄로 이어져야 한다(§2 «끊기는 자리»)."""
    ses = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "edit" / "session.py").read_text(encoding="utf-8")
    assert 'data["edge_len_mm"]' in ses
    res = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "restrict.py").read_text(encoding="utf-8")
    assert res.count("edge_len_mm=") == 2, "두 호출부 중 하나가 빠졌다"
    assert '"declared_pipes"' in res


def test_표가_접힌_배관을_표시한다():
    """§5 기준 8 — 좌표↔length 검사가 그 부류를 알아보려면 행에 남아야 한다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "tables.py").read_text(encoding="utf-8")
    assert 'row["len_src"] = "신축배관 접기"' in src
    api = (_ROOT / "routes" / "module_f" / "api_design.py").read_text(
        encoding="utf-8")
    assert "declared_pipes=got.get(" in api


# ─────────────────────────────── ⑤ 지시서 §6 금지사항
def test_직교화로_펴지_않는다():
    """없던 90° 엘보를 만들고, 가지관의 실제 대각 주행을 망가뜨린다."""
    src = (_ROOT / "routes" / "module_f" / "flexfold.py").read_text(
        encoding="utf-8")
    assert "orthogonalize" not in src


def test_좌표_계산식은_안_건드린다():
    """`bake_isometric` · `normalize_node_coords` 는 이 작업 범위 밖이다."""
    src = (_ROOT / "routes" / "module_f" / "flexfold.py").read_text(
        encoding="utf-8")
    for name in ("bake_isometric", "normalize_node_coords",
                 "_head_attachments"):
        assert name not in src, name
