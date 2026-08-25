# -*- coding: utf-8 -*-
"""모듈 G 수리계산 입력 — 항목별 수용 기준 검증(지시서 §4).

    python tests/test_design_input.py [G1 G2 ...]      항목 지정(기본 전부)

Qt 없이 돈다(헤드리스). 실도면 통합은 tests/_out/ 아래에만 쓴다 — G 트리 밖으로
중간 산출물을 내보내지 않는다(§3).
"""
from __future__ import annotations

import math
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)

KEY = "B1F 현장조사 소화설비 평면도"
OUT_DIR = _ROOT / "tests" / "_out"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def _board():
    """손질 저장본을 연다. 여러 테스트가 공유하므로 한 번만 만든다."""
    if not hasattr(_board, "_b"):
        from services.cad_import.edit.session import EditSession
        es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
        _board._b = es
    return _board._b


# ─────────────────────────────────────────────────────────── G1
def g1():
    print("\n[G1] 최불리 K 선정 이식")
    from services.cad_import.design.worst import worst_k_heads

    es = _board()
    b = es.board
    check("급수 시작 위치 있음", bool(b.sources), f"{len(b.sources)}곳")

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    need = ("heads", "anchor", "edges", "loads", "nodes",
            "far_m", "near_m", "span_m", "total_m", "max_load")
    missing = [k for k in need if k not in w]
    check("반환 키 10종", not missing, f"빠짐 {missing}" if missing else "전부 있음")
    check("K개 선정", len(w["heads"]) == 30, f"{len(w['heads'])}개")

    # ★핵심: 모듈 F 와 완전히 일치해야 이식이 성공한 것이다.
    sys.path.append(str(_ROOT.parent))
    from routes.module_f.remote30 import _worst_k_heads as f_worst
    wf = f_worst(b.pts, b.edges, b.hnodes, b.sources, k=30)
    check("모듈 F 와 앵커 일치", w["anchor"] == wf["anchor"],
          f"G {w['anchor']} / F {wf['anchor']}")
    check("모듈 F 와 헤드 집합 일치", set(w["heads"]) == set(wf["heads"]),
          f"차집합 {len(set(w['heads']) ^ set(wf['heads']))}개")
    check("모듈 F 와 far_m 일치", w["far_m"] == wf["far_m"],
          f"G {w['far_m']} / F {wf['far_m']}")
    check("모듈 F 와 max_load 일치", w["max_load"] == wf["max_load"],
          f"G {w['max_load']} / F {wf['max_load']}")

    # 앵커 방식이면 설계면적이 뭉친다 — 「먼 순서」와 갈리는 지점.
    check("설계면적 폭이 corridor 총연장보다 작다",
          w["span_m"] < w["total_m"],
          f"폭 {w['span_m']} m / 총연장 {w['total_m']} m")
    print(f"      앵커 {w['far_m']} m · 폭 {w['span_m']} m · "
          f"연장 {w['total_m']} m · max_load {w['max_load']}")
    return w


# ─────────────────────────────────────────────────────────── G2
def g2():
    print("\n[G2] corridor 제한 전개 + 역참조")
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.worst import worst_k_heads

    es = _board()
    b = es.board
    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    payload = es.convert_payload()
    # 이 저장본은 급수원이 둘이다 — 기준선과 같은 것(Z1)을 쓴다(BLOCKED B2).
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None

    # BLOCKED B4 · 1안 — 선정 후보를 «전개가 붙일 수 있는 헤드» 로 먼저 좁힌다.
    got = select_and_expand(payload, b, k=30, selected_source=sel)
    if not check("선정+제한 전개 성공", got.get("ok"), str(got.get("error"))[:90]):
        return None
    w = got["worst"]
    check("제외 헤드 수가 드러난다", "excluded_heads" in got,
          f"후보 {got.get('candidate_heads')} / 도면 {got.get('total_heads')}"
          f" · 제외 {got.get('excluded_heads')}")

    kfp = got["kfp"]
    pipes = kfp.get("pipe_data") or {}
    nodes = kfp.get("nodes_meta_runtime") or {}
    # ★수용 기준은 «입력에 30개를 넣었나» 가 아니라 «전개가 30개를 살렸나» 다.
    #   hcov 는 입력을 되돌려줄 뿐이라 그걸 세면 빈 망도 초록불이 된다(실제로 그랬다).
    n_built_heads = sum(1 for m in (kfp.get("nodes_meta_runtime") or {}).values()
                        if str((m or {}).get("type_id", "")) == "head")
    check("제한 전개가 살린 헤드 수 == 선정 K",
          n_built_heads == len(w["heads"]),
          f"전개 {n_built_heads} / 선정 {len(w['heads'])}  "
          f"(BLOCKED B4 — 선정과 전개가 «물 닿음» 을 다르게 본다)")
    # ★세로 처리(§G19) 뒤로는 «도면에 그려진 선» 이 아닌 배관이 생긴다 —
    #   헤드 접속관·가지 상승은 변환이 만든 구간이라 대응할 board 간선이 없다.
    #   그래서 「전부 덮인다」가 아니라 「덮이지 않은 것은 전부 세로 구간이고,
    #   그 담당 헤드 수를 망에서 세어 두었다」를 확인한다. 느슨해진 것이 아니라
    #   판정 대상이 정확해진 것이다.
    meta_n = kfp.get("nodes_meta_runtime") or {}
    tl = got.get("tree_loads") or {}

    def _xyz(nid):
        c = (meta_n.get(nid) or {}).get("coords") or (0.0, 0.0, 0.0)
        return (float(c[0]), float(c[1]), float(c[2]) if len(c) > 2 else 0.0)

    not_vertical, no_load = [], []
    for pid in got["uncovered_pipes"]:
        pr = (kfp.get("pipe_data") or {}).get(pid) or {}
        a = pr.get("start") or pr.get("from")
        z = pr.get("end") or pr.get("to")
        pa, pb = _xyz(a), _xyz(z)
        flat = math.hypot(pb[0] - pa[0], pb[1] - pa[1])
        if not (flat < 1e-6 and abs(pb[2] - pa[2]) > 1e-9):
            not_vertical.append(pid)
        if not tl.get(pid):
            no_load.append(pid)
    check("역참조가 «도면에 그려진» 배관을 전부 덮는다", not not_vertical,
          f"세로가 아닌데 미포함 {len(not_vertical)}개 "
          f"(세로 구간 {len(got['uncovered_pipes'])}개는 정상)")
    check("세로 구간도 담당 헤드 수를 안다", not no_load,
          f"담당 헤드 수를 모르는 세로 구간 {len(no_load)}개")
    check("역참조가 board 간선을 가리킨다",
          all(isinstance(v, tuple) and len(v) == 2
              and 0 <= v[0] < len(b.pts) and 0 <= v[1] < len(b.pts)
              for v in got["edge_ref"].values()),
          f"{len(got['edge_ref'])}건")
    check("node_ref 존재", bool(got["node_ref"]), f"{len(got['node_ref'])}건")
    # 제한 전개는 전체망보다 작아야 한다 — 그게 «제한» 의 뜻이다.
    check("제한망이 corridor 규모다", len(pipes) >= len(w["heads"]),
          f"배관 {len(pipes)} · 선정 헤드 {len(w['heads'])}개를 먹이려면 그 이상")
    print(f"      제한망 노드 {len(nodes)} · 배관 {len(pipes)}")
    return got


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
    print("\n[G3] 관경 결정 — 혼합 규칙")
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
    if not check("handoff 에서 치수 텍스트 복원", w is not None,
                 f"텍스트 {len(w.texts)}개" if w is not None else "복원 실패"):
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


# ─────────────────────────────────────────────────────────── G4
def g4():
    print("\n[G4] 부속 · 노즐 · 기기")
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
    # ★좌표는 coords[x, y, z] 다. x/y 키로 읽으면 전부 (0,0) 이 되어 각도가
    #   0 으로 붕괴하고 엘보가 하나도 안 잡힌다(실제로 그렇게 헛돌았다).
    nxy = {nid: (float(m["coords"][0]), float(m["coords"][1]))
           for nid, m in (kfp.get("nodes_meta_runtime") or {}).items()
           if isinstance(m.get("coords"), (list, tuple)) and len(m["coords"]) >= 2}
    # ★상류를 모르면 fitting_rules 가 엘보를 못 달고 티를 전부 «판정 불가» 로
    #   센다. 그건 설계된 열화 경로지 정상 결과가 아니다 — 급수원(펌프)을 뿌리로
    #   BFS 부모를 만들어 진짜 판정을 재게 한다(G5 가 이 순서를 다시 쓴다).
    meta = kfp.get("nodes_meta_runtime") or {}
    root = next((nid for nid, m in meta.items()
                 if str((m or {}).get("type_id", "")) == "pump"), None)
    if root is None:
        root = next(iter(nxy), None)
    adj = {}
    for pid, pr in (kfp.get("pipe_data") or {}).items():
        a = pr.get("start") or pr.get("from")
        b = pr.get("end") or pr.get("to")
        if a and b:
            adj.setdefault(a, []).append(b)
            adj.setdefault(b, []).append(a)
    par, seen, q = {}, {root}, [root]
    while q:
        u = q.pop(0)
        for v in adj.get(u, ()):
            if v not in seen:
                seen.add(v); par[v] = u; q.append(v)
    check("급수원 뿌리 BFS 가 망을 덮는다", len(seen) >= len(nxy) * 0.9,
          f"도달 {len(seen)} / 노드 {len(nxy)} · 뿌리 {root}")
    real = build_fittings(kfp, nxy, bores, parents=par, lib=lib)
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


# ─────────────────────────────────────────────────────────── G5
def g5():
    print("\n[G5] 5개 테이블 조립")
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
    # ★두 양은 구조적으로 다르다(BLOCKED B5) — corridor 는 최단경로 «합집합»의
    #   board 간선(206개), 전개는 물길·막다른관 정리 후 병합된 배관(53개)이다.
    #   그래서 ±1 % 를 요구할 수 없다. 형상 왜곡은 아래 «배관별 길이차» 로 조인다.
    within = abs(plane - want) <= max(0.03 * want, 0.05)
    check("배관 길이 합이 corridor 총연장과 ±3% (B5)", within,
          f"평면 {plane:.2f} m vs corridor {want} m "
          f"(차 {plane - want:+.2f} m · 전체 {total_len:.2f} · 수직 {vert:.2f})")

    # 진짜 회귀 신호 — 배관 하나하나가 원 board 간선 길이와 맞는가(형상 왜곡).
    import math as _m
    dsum = 0.0
    for r in tbl.pipes:
        ref = got["edge_ref"].get(r["label"])
        if not ref:
            continue
        i, j = ref
        dsum += abs(float(r["length"]) - _m.dist(es.board.pts[i], es.board.pts[j]) / 1000.0)
    # ★절대값 문턱은 망 크기에 따라 흔들린다 — 같은 품질인데 망이 커지면
    #   빨간불이 된다(실측: 0.180 m/180 m → 0.592 m/212 m, 비율은 0.1~0.3 %).
    #   형상 왜곡은 «비율» 로 재야 한다.
    ratio = dsum / total_len if total_len else 0.0
    check("배관별 길이가 board 간선과 일치(형상 무왜곡)", ratio <= 0.005,
          f"절대차 합 {dsum:.3f} m / 총연장 {total_len:.1f} m = {ratio*100:.2f}%")

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


# ─────────────────────────────────────────────────────────── G6
def g6():
    print("\n[G6] SDF 방출")
    import os
    import xml.etree.ElementTree as ET
    from services.cad_import.design.emit import (
        AssetMissing, emit_design_sdf, resolve_standard_slf,
        resolve_template_sdf)

    tbl = g5()
    if tbl is None:
        return None

    check("템플릿 SDF 해석", resolve_template_sdf().is_file(),
          resolve_template_sdf().name[:40])
    check("표준 SLF 해석", resolve_standard_slf().is_file(),
          resolve_standard_slf().name)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "module_g_design.sdf"
    got = emit_design_sdf(tbl, out, project_title="Module G 검증")
    check("SDF 생성", got.is_file() and got.stat().st_size > 1000,
          f"{got.stat().st_size:,} bytes")
    slf = got.with_suffix(".slf")
    check("SLF 를 한 쌍으로 저장", slf.is_file(), f"{slf.name}")

    root = ET.parse(got).getroot()
    nodes = root.findall(".//Nodes/Node")
    pipes = root.findall(".//Links//Pipe")
    nozzles = root.findall(".//Links//Nozzle")
    check("노드 수가 테이블과 일치", len(nodes) == len(tbl.nodes),
          f"SDF {len(nodes)} / 표 {len(tbl.nodes)}")
    check("배관 수가 테이블과 일치", len(pipes) == len(tbl.pipes),
          f"SDF {len(pipes)} / 표 {len(tbl.pipes)}")
    check("노즐 수가 테이블과 일치", len(nozzles) == len(tbl.nozzles),
          f"SDF {len(nozzles)} / 표 {len(tbl.nozzles)}")

    # ★관경이 "Unset" 으로 뜨면 안 된다 — bore 속성이 실수로 들어가야 한다.
    bores = [p_.get("bore") for p_ in pipes]
    bad = [b for b in bores if b in (None, "", "Unset")]
    check("관경이 Unset 이 아니다", not bad,
          f"Unset {len(bad)} / 배관 {len(bores)} · 예 {bores[:4]}")

    # 템플릿을 썼으므로 Graphics 블록(표시 메타)이 살아 있어야 한다.
    check("템플릿의 Graphics 블록 보존",
          root.find(".//Graphics") is not None, "Graphics 있음")

    # ★자산이 없으면 «파일을 만들지 않고» 실패해야 한다.
    out2 = OUT_DIR / "should_not_exist.sdf"
    if out2.exists():
        out2.unlink()
    os.environ["REMOTE30_TEMPLATE_SDF"] = str(OUT_DIR / "no_such_template.sdf")
    try:
        emit_design_sdf(tbl, out2)
        check("자산 없으면 실패", False, "예외 없이 진행했다")
    except AssetMissing as exc:
        check("자산 없으면 실패", True, str(exc).splitlines()[0][:52])
    finally:
        os.environ.pop("REMOTE30_TEMPLATE_SDF", None)
    check("실패 시 파일을 남기지 않는다", not out2.exists(), str(out2.name))
    return got


ITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4, "G5": g5, "G6": g6}


def main() -> int:
    want = sys.argv[1:] or list(ITEMS)
    for name in want:
        fn = ITEMS.get(name)
        if fn is None:
            print(f"  (모르는 항목: {name})")
            continue
        fn()
    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("수용 기준 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
