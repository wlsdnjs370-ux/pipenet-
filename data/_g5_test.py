# -*- coding: utf-8 -*-
"""[G5] 수용 기준 검사 추가."""
import io

p = "cad_project_editor_g/tests/test_design_input.py"
s = io.open(p, encoding="utf-8").read()

block = '''

# ─────────────────────────────────────────────────────────── G5
def g5():
    print("\\n[G5] 5개 테이블 조립")
    from services.cad_import.design.bore import extract_dia_text_points
    from services.cad_import.design.tables import build_design_tables

    got = g2()
    if got is None:
        return None
    es = _board()
    w = _world()
    tbl = build_design_tables(
        got["kfp"], got["worst"], got["edge_ref"],
        extract_dia_text_points(w.texts),
        board_pts=es.board.pts,
        excluded_heads=got.get("excluded_heads", 0))

    node_labels = {r["label"] for r in tbl.nodes}
    orphan = [r["label"] for r in tbl.pipes
              if r["in"] not in node_labels or r["out"] not in node_labels]
    check("배관 in/out 이 전부 노드표에 있다(고아 0)", not orphan,
          f"고아 {len(orphan)}건 {orphan[:4]}")

    fit_orphan = [f["pipe"] for f in tbl.fittings
                  if f["pipe"] not in {r["label"] for r in tbl.pipes}]
    check("부속표의 배관이 전부 배관표에 있다", not fit_orphan,
          f"고아 {len(fit_orphan)}건")

    check("노즐 수 == 선정 K",
          len(tbl.nozzles) == len(got["worst"]["heads"]),
          f"노즐 {len(tbl.nozzles)} / 선정 {len(got['worst']['heads'])}")

    # 길이 합 — 수직 전개분을 뺀 평면 성분 기준으로 ±1 %
    total_len = sum(float(r["length"]) for r in tbl.pipes)
    vert = sum(abs(float(r["elev"])) for r in tbl.pipes)
    plane = total_len - vert
    want = float(got["worst"]["total_m"])
    within = abs(plane - want) <= max(0.01 * want, 0.05)
    check("배관 길이 합이 corridor 총연장과 ±1%", within,
          f"평면 {plane:.2f} m vs corridor {want} m "
          f"(전체 {total_len:.2f} · 수직 {vert:.2f})")

    # 표 첫 행 = 급수원 인접 배관 / 트리 꼬리 = 루프 잔여
    first = tbl.pipes[0] if tbl.pipes else {}
    root_lab = next((r["label"] for r in tbl.nodes
                     if r.get("io_node") == "Input"), None)
    check("첫 배관이 급수원에 붙어 있다",
          root_lab is not None and first.get("in") == root_lab,
          f"뿌리 {root_lab} · 첫 배관 in={first.get('in')}")
    off = [r for r in tbl.pipes if r.get("off_tree")]
    check("루프 잔여는 표 꼬리에 몰린다",
          not off or all(r.get("off_tree") for r in tbl.pipes[-len(off):]),
          f"루프 잔여 {len(off)}건")

    # 단위(§T3) — 노드 좌표만 mm
    xs = [abs(r["x"]) for r in tbl.nodes if r["x"]]
    check("노드 좌표가 mm 자리수", bool(xs) and max(xs) > 1000,
          f"|x| 최대 {max(xs) if xs else 0}")

    keys = {"label", "in", "out", "type", "dia", "length", "elev",
            "c", "status", "group"}
    check("배관 행 키가 PipeTables 규약", keys <= set(tbl.pipes[0]),
          f"빠짐 {sorted(keys - set(tbl.pipes[0]))}")
    print(f"      노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} · "
          f"노즐 {len(tbl.nozzles)} · 부속 {len(tbl.fittings)} · "
          f"기기 {len(tbl.equipment)} · meta {len(tbl.meta)}행")
    return tbl


ITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4, "G5": g5}'''

s = s.replace('\n\nITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4}', block, 1)
assert '"G5": g5' in s
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("G5 검사 추가")
