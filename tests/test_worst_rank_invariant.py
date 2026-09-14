# -*- coding: utf-8 -*-
"""[최불리규칙복원 §3] 「먼 순서 그대로 K 개」를 식으로 못박는다.

지시서 `ModuleF_최불리규칙_복원_지시서.md`.

■ 규칙 (§0)

    자동 추출 + 사람의 손질로 «정의된» 헤드 **전부**를 후보로 두고, 급수원에서
    배관을 따라 잰 유하거리가 **긴 순서 그대로** 기준개수 K 개를 뽑는다.
    영역·도면 장을 지정했으면 그 안으로만 가둔다. **그 밖의 이유로 줄이지
    않는다.**

■ 식 (§3)

    C = 후보 (영역·장 제한을 통과하고 급수원에서 도달하는 헤드 · 자리 대표)
    S = 선정 결과

    ①  |S| = K
    ②  min{ far(h) : h ∈ S }  ≥  max{ far(h) : h ∈ C − S }
    ③  S ⊆ C

■ 왜 시험으로 못박나

  말로 적힌 규칙은 깨졌다. 「전개가 못 붙이는 헤드」를 후보에서 빼는 조치가
  들어갔고, `worst_k_heads` 는 `only_heads` 밖을 순위 계산에서 아예 건너뛴다
  (`continue`). 가장 먼 헤드는 말단 가지관 끝에 있어 «관이 스쳐 지나가는»
  모양이 되기 쉬운데, 하필 그것들이 빠졌다.

      실측(대명동 K=30 · `scripts/_probe_candidate_drop.py`):
        빠진 33개 중 **상위 30 안에 5개** — 2등 51.69m · 5등 49.49m ·
        19등 · 24등 · 26등. 전부 `pass_under`.
        그 자리에 31·32·33·35·36등이 들어왔다.

  그래서 ②가 **깨지는 모습 자체**를 아래 회귀 시험에 박아 둔다.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.worst import worst_k_heads   # noqa: E402


def _src(rel):
    return (_ROOT / rel).read_text(encoding="utf-8")


# ═══════════════════════ 인공 망 — 급수원에서 한 줄로 뻗은 가지관
#
#   노드 0 이 급수원, 노드 i 가 i·1000 mm 자리. 헤드 h(i) 는 노드 i+1 에 붙는다.
#   그래서 far(h(i)) = (i+1)·1000 mm — 순위가 자명해 «먼 순서» 를 눈으로 검산할
#   수 있다. 좌표를 다르게 줘야 `_place` 가 «같은 자리» 로 접지 않는다.
N = 12


def _net(n=N):
    pts = [(float(i * 1000), 0.0) for i in range(n + 1)]
    edges = {(i, i + 1) for i in range(n)}
    hnodes = [{i + 1} for i in range(n)]
    head_xy = [(float((i + 1) * 1000), 0.0) for i in range(n)]
    return pts, edges, hnodes, [0], head_xy


def _pick(k, only=None):
    pts, edges, hnodes, src, hxy = _net()
    return worst_k_heads(pts, edges, hnodes, src, k=k,
                         only_heads=only, source_index=0, head_xy=hxy)


def _far_all(only=None):
    """후보 전부의 유하거리 — k 를 후보 수만큼 줘서 전 순위를 받는다."""
    w = _pick(N if only is None else max(1, len(only)), only=only)
    return {int(h): float(v) for h, v in (w.get("dists") or {}).items()}


def _check(S, C, far, k):
    """①②③ — 깨진 것을 문자열로 돌려준다(빈 목록이면 성립)."""
    bad = []
    if len(S) != k:
        bad.append(f"① |S|={len(S)} ≠ K={k}")
    if not set(S) <= set(C):
        bad.append(f"③ 후보 밖 {sorted(set(S) - set(C))}")
    rest = [h for h in C if h not in set(S)]
    if S and rest:
        mn = min(far[h] for h in S)
        mx = max(far[h] for h in rest)
        if mn + 0.1 < mx:
            bad.append(f"② 뽑힌 꼴찌 {mn} < 안 뽑힌 1등 {mx}")
    return bad


# ═══════════════════════ ①②③ — 후보를 안 깎으면 선다
def test_먼_순서_그대로_K개():
    far = _far_all()
    C = sorted(far)
    for k in (1, 3, 7, N):
        w = _pick(k)
        assert not _check(list(w["heads"]), C, far, k), (k, w["heads"])


def test_가장_먼_헤드가_반드시_든다():
    """★1등이 빠지면 그것이 곧 증상이다 — 최원 유하거리가 통째로 달라진다."""
    far = _far_all()
    top = max(far, key=lambda h: far[h])
    for k in (1, 3, 7):
        assert top in _pick(k)["heads"], k


# ═══════════════════════ 영역으로 가두면 «그 안에서» 선다
def test_영역_안에서도_먼_순서다():
    """§0 — 영역·장 제한은 **허용된 유일한** 후보 제한이다."""
    zone = {0, 1, 2, 3, 4}            # 가까운 쪽 다섯만
    far = _far_all(zone)
    C = sorted(far)
    for k in (1, 2, 5):
        w = _pick(k, only=zone)
        assert not _check(list(w["heads"]), C, far, k), (k, w["heads"])
    # 영역 밖의 더 먼 헤드는 **안 들어온다** — 가두는 것이 목적이다.
    assert set(_pick(2, only=zone)["heads"]) <= zone


# ═══════════════════════ ★회귀 — 「붙는 헤드」로 좁히면 ②가 깨진다
def test_후보를_깎으면_먼_순서가_깨진다():
    """★지금 증상을 **그 모습 그대로** 박아 둔다.

    `worst_k_heads` 는 `only_heads` 밖을 순위 계산에서 건너뛴다. 그래서 먼
    헤드를 「못 붙는다」고 빼면 그만큼 **가까운 헤드가 K 안으로 올라온다.**
    이 시험이 깨지는 날은 누군가 그 좁히기를 되살린 날이다.
    """
    far = _far_all()
    C = sorted(far)
    k = 3
    top2 = sorted(far, key=lambda h: -far[h])[:2]       # 1·2등을 «못 붙는다» 로
    keep = set(C) - set(top2)

    narrowed = list(_pick(k, only=keep)["heads"])
    assert not (set(top2) & set(narrowed)), "좁혔는데도 들어왔다 — 전제가 틀림"
    # 후보 전체로 재면 ②가 깨진다 — 이것이 증상이다.
    bad = _check(narrowed, C, far, k)
    assert any(v.startswith("②") for v in bad), bad

    # 안 좁히면 같은 K 로 ②가 선다.
    assert not _check(list(_pick(k)["heads"]), C, far, k)


# ═══════════════════════ 제품 코드가 그 좁히기를 안 한다 (§2-1)
def test_손질이_후보를_안_깎는다():
    s = _src("routes/module_f/api_edit.py")
    i = s.index("def _compute_worst")
    seg = s[i:s.index("w = _worst_k_heads(", i)]
    # ★주석은 걷고 본다 — 「왜 없앴는지」를 적어 두면 그 설명 안에 옛 코드가
    #   그대로 들어간다. 시험이 설명을 읽고 깨지면 설명을 못 적게 된다.
    code = "\n".join(ln for ln in seg.splitlines()
                     if not ln.lstrip().startswith("#"))
    # 후보를 keep 으로 갈아 끼우는 **줄** 이 없어야 한다.
    for bad in ("only = before & keep", "only = cand & keep",
                "only = (before & keep)", "only = (cand & keep)"):
        assert bad not in code, f"후보 좁히기가 살아 있다: {bad}"
    # ★고르기 **전에** 전체망을 전개하지 않는다 (2026-09-14 속도).
    #   B1F 실측으로 그 한 줄이 154.10초 · 98.2% 였다. 「붙는가」는 고른 뒤
    #   그 K 개만 본다 — 답은 같고 0.91초다(194배).
    assert "wet_heads(" not in code, "고르기 전에 전체망을 전개한다"


def test_선정_인자를_안_바꿨다():
    """§6 — `hnodes` 를 다른 목록으로 갈아 끼우면 유하거리가 통째로 달라진다."""
    s = _src("routes/module_f/api_edit.py")
    i = s.index("w = _worst_k_heads(")
    seg = s[i:i + 300]
    assert "b.pts, b.edges, b.hnodes, b.sources" in seg
    assert "head_centers" not in seg


def test_매_요청마다_검사하고_응답에_싣는다():
    s = _src("routes/module_f/api_edit.py")
    assert "def _rank_invariant(" in s
    assert "rank_inv = _rank_invariant(" in s
    assert '"rank_invariant": rank_inv' in s, "응답에 안 싣는다"
    i = s.index("def _rank_invariant(")
    seg = s[i:s.index("\ndef ", i + 10)]
    for mark in ("①", "②", "③"):
        assert mark in seg, mark
