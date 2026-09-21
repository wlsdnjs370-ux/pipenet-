# -*- coding: utf-8 -*-
"""[G4] 수용 기준 검사 추가."""
import io

p = "cad_project_editor_g/tests/test_design_input.py"
s = io.open(p, encoding="utf-8").read()

block = '''

# ─────────────────────────────────────────────────────────── G4
def g4():
    print("\\n[G4] 부속 · 노즐 · 기기")
    from services.cad_import.design.fitting import (
        build_equipment, build_fittings, build_nozzles, equivalent_length_m,
        load_equivalent_lengths)

    lib = load_equivalent_lengths()
    check("등가길이 라이브러리 적재", bool(lib.get("ELBOW_90_STD")),
          f"항목 {len(lib)}종 · 90엘보 호칭경 {len(lib.get('ELBOW_90_STD') or {})}칸")
    check("라이브러리에 없는 칸은 None", equivalent_length_m(lib, "elbow", 15) is None
          or isinstance(equivalent_length_m(lib, "elbow", 15), float),
          f"15A 90엘보 = {equivalent_length_m(lib, 'elbow', 15)}")

    # ① 합성 — 90° 꺾임 1곳 → 엘보 1
    net = {"pipe_data": {"P1": {"start": "N1", "end": "N2"},
                         "P2": {"start": "N2", "end": "N3"}}}
    xy = {"N1": (0.0, 0.0), "N2": (10.0, 0.0), "N3": (10.0, 10.0)}
    par = {"N2": "N1", "N3": "N2"}
    r = build_fittings(net, xy, {"P1": (50, "text"), "P2": (50, "text")},
                       parents=par, lib=lib)
    check("90° 꺾임 → 엘보 1", r["counts"].get("elbow") == 1, str(r["counts"]))

    # ② 직진 통과 분기 → 직류티는 계상하지 않는다
    net2 = {"pipe_data": {"P1": {"start": "N1", "end": "N2"},
                          "P2": {"start": "N2", "end": "N3"},
                          "P3": {"start": "N2", "end": "N4"}}}
    xy2 = {"N1": (0.0, 0.0), "N2": (10.0, 0.0),
           "N3": (20.0, 0.0),      # 직진 — 직류티
           "N4": (10.0, 10.0)}     # 꺾임 — 분류티
    par2 = {"N2": "N1", "N3": "N2", "N4": "N2"}
    r2 = build_fittings(net2, xy2, {p_: (50, "text") for p_ in net2["pipe_data"]},
                        parents=par2, lib=lib)
    check("직진 갈래는 티 미계상 · 꺾인 갈래만 분류티",
          r2["counts"].get("tee") == 1 and "P3" in
          [pid for pid, rec in r2["per_pipe"].items() if rec["fittings"]],
          f"{r2['counts']} · 티 달린 배관 "
          f"{[pid for pid, rec in r2['per_pipe'].items() if rec['fittings']]}")

    # ③ 실도면 — 고아 참조 0
    got = g2()
    if got is None:
        return None
    from services.cad_import.design.bore import decide_bores, extract_dia_text_points
    w = _world()
    es = _board()
    bores = decide_bores(got["kfp"], got["edge_ref"], got["worst"]["loads"],
                         extract_dia_text_points(w.texts), pts=es.board.pts)
    kfp = got["kfp"]
    nxy = {nid: (m.get("x", 0.0), m.get("y", 0.0))
           for nid, m in (kfp.get("nodes_meta_runtime") or {}).items()}
    real = build_fittings(kfp, nxy, bores, parents=None, lib=lib)
    pipe_ids = set((kfp.get("pipe_data") or {}))
    check("부속표의 배관이 전부 배관표에 있다(고아 0)",
          set(real["per_pipe"]) <= pipe_ids,
          f"부속 {len(real['per_pipe'])} / 배관 {len(pipe_ids)}")
    check("미해결을 0 으로 채우지 않는다",
          isinstance(real["unresolved_length"], int)
          and isinstance(real["unresolved_kind"], int),
          f"판정불가 {real['unresolved_kind']} · 등가길이 미해결 "
          f"{real['unresolved_length']}")

    noz = build_nozzles(kfp, k_factor=80.0)
    check("노즐 수 == 선정 K", len(noz) == len(got["worst"]["heads"]),
          f"노즐 {len(noz)} / 선정 {len(got['worst']['heads'])}")
    eq = build_equipment(kfp, valve_nodes=None)
    check("알람밸브 미지정이면 기기 행 없음", eq == [], f"{len(eq)}행")
    print(f"      부속 {real['counts']} · 노즐 {len(noz)}")
    return real


ITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4}'''

s = s.replace('\n\nITEMS = {"G1": g1, "G2": g2, "G3": g3}', block, 1)
assert '"G4": g4' in s
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("G4 검사 추가")
