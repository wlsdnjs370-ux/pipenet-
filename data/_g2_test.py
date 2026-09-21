# -*- coding: utf-8 -*-
"""[G2] 수용 기준 검사를 tests/test_design_input.py 에 추가."""
import io

p = "cad_project_editor_g/tests/test_design_input.py"
s = io.open(p, encoding="utf-8").read()

block = '''

# ─────────────────────────────────────────────────────────── G2
def g2():
    print("\\n[G2] corridor 제한 전개 + 역참조")
    from services.cad_import.design.restrict import expand_worst
    from services.cad_import.design.worst import worst_k_heads

    es = _board()
    b = es.board
    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    payload = es.convert_payload()
    # 이 저장본은 급수원이 둘이다 — 기준선과 같은 것(Z1)을 쓴다(BLOCKED B2).
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None

    got = expand_worst(payload, b, w, selected_source=sel)
    if not check("제한 전개 성공", got.get("ok"), str(got.get("error"))[:90]):
        return None

    kfp = got["kfp"]
    pipes = kfp.get("pipe_data") or {}
    nodes = kfp.get("nodes_meta_runtime") or {}
    check("헤드 수 == 선정 K", len(got["hcov"] or []) == len(w["heads"]),
          f"{len(got['hcov'] or [])} / {len(w['heads'])}")
    check("역참조가 모든 배관을 덮는다", not got["uncovered_pipes"],
          f"미포함 {len(got['uncovered_pipes'])}개")
    check("역참조가 board 간선을 가리킨다",
          all(isinstance(v, tuple) and len(v) == 2
              and 0 <= v[0] < len(b.pts) and 0 <= v[1] < len(b.pts)
              for v in got["edge_ref"].values()),
          f"{len(got['edge_ref'])}건")
    check("node_ref 존재", bool(got["node_ref"]), f"{len(got['node_ref'])}건")
    # 제한 전개는 전체망보다 작아야 한다 — 그게 «제한» 의 뜻이다.
    check("전체망보다 작다", len(pipes) < 1444,
          f"제한 {len(pipes)} 배관 < 전체 1444")
    print(f"      제한망 노드 {len(nodes)} · 배관 {len(pipes)}")
    return got


ITEMS = {"G1": g1, "G2": g2}'''

s = s.replace('\n\nITEMS = {"G1": g1}', block, 1)
assert "ITEMS = {\"G1\": g1, \"G2\": g2}" in s
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("G2 검사 추가")
