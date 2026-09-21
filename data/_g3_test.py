# -*- coding: utf-8 -*-
"""[G3] 수용 기준 검사 추가."""
import io

p = "cad_project_editor_g/tests/test_design_input.py"
s = io.open(p, encoding="utf-8").read()

block = '''

# ─────────────────────────────────────────────────────────── G3
def _world():
    """치수 텍스트 원본. handoff 캐시에 이미 있다 — DXF 를 다시 읽지 않는다."""
    if not hasattr(_world, "_w"):
        import json
        from services.cad_import.pipeline import handoff, stage1 as s1
        spec_path = (_ROOT / "docs" / "import" / "0단계_새찍기"
                     / f"{KEY}_찍은스펙.json")
        src = json.loads(spec_path.read_text(encoding="utf-8")).get("source_dxf")
        _world._w = handoff.load_world(KEY, src, s1.World)
    return _world._w


def g3():
    print("\\n[G3] 관경 결정 — 혼합 규칙")
    from services.cad_import.design.bore import (
        decide_bores, extract_dia_text_points, nfpc_min_bore_mm, source_counts)

    # ① 별표1 매핑 — 지시서에 적힌 값 그대로
    table = [(1, 25), (3, 32), (5, 40), (10, 50), (30, 65),
             (60, 80), (100, 100), (200, 150)]
    bad = [(n, nfpc_min_bore_mm(n), want) for n, want in table
           if nfpc_min_bore_mm(n) != want]
    check("별표1 매핑 8종", not bad, str(bad) if bad else "1→25 … 200→150")

    # ② 안전측 — 텍스트 50A 가 1200mm 거리, 별표1 최소 65 → 65 · nfpc_min
    net = {"pipe_data": {"P1": {}}}
    edge_ref = {"P1": (0, 1)}
    pts = [(0.0, 0.0), (10000.0, 0.0)]
    loads = {(0, 1): 30}                      # 30개 담당 → 별표1 65
    texts = [(5000.0, 1200.0, 50)]            # 1200mm 떨어진 "50A"
    got = decide_bores(net, edge_ref, loads, texts, pts=pts)
    check("텍스트<별표1 이면 별표1 채택", got["P1"] == (65, "nfpc_min"),
          f"{got['P1']} (기대 (65, 'nfpc_min'))")

    # 텍스트가 별표1 보다 크면 텍스트를 쓴다
    got2 = decide_bores(net, edge_ref, {(0, 1): 3},
                        [(5000.0, 300.0, 100)], pts=pts)
    check("텍스트>별표1 이면 텍스트 채택", got2["P1"] == (100, "text"),
          f"{got2['P1']}")
    # 1500mm 밖 텍스트는 안 잡힌다 → fallback
    got3 = decide_bores(net, edge_ref, {(0, 1): 3},
                        [(5000.0, 2000.0, 100)], pts=pts)
    check("1500mm 밖 텍스트는 무시", got3["P1"] == (32, "nfpc_fallback"),
          f"{got3['P1']}")

    # ③ 실도면 — text 비율이 0% 면 어댑터가 죽은 것이다
    w = _world()
    if not check("handoff 에서 치수 텍스트 복원", w is not None, "world None"):
        return None
    dia_pts = extract_dia_text_points(w.texts)
    check("관경 텍스트 추출", len(dia_pts) > 0, f"{len(dia_pts)}개 / 텍스트 {len(w.texts)}")

    got_g2 = g2()
    if got_g2 is None:
        return None
    es = _board()
    bores = decide_bores(got_g2["kfp"], got_g2["edge_ref"],
                         got_g2["worst"]["loads"], dia_pts, pts=es.board.pts)
    cnt = source_counts(bores)
    check("모든 배관에 관경이 붙는다",
          len(bores) == len((got_g2["kfp"].get("pipe_data") or {})),
          f"{len(bores)}건")
    check("실도면에서 text 비율이 0% 가 아니다", cnt["text"] > 0,
          f"text {cnt['text']} · nfpc_min {cnt['nfpc_min']} · "
          f"fallback {cnt['nfpc_fallback']}")
    dias = sorted({d for d, _s in bores.values()})
    print(f"      관경 분포 {dias} · 근거 {cnt}")
    return bores


ITEMS = {"G1": g1, "G2": g2, "G3": g3}'''

s = s.replace('\n\nITEMS = {"G1": g1, "G2": g2}', block, 1)
assert '"G3": g3' in s
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("G3 검사 추가")
