# -*- coding: utf-8 -*-
"""[G8] 모듈 A 대조 — 같은 도면을 A 로도 돌려 위상이 같은지 본다.

지시서 §G8 의 다섯 항목:
    앵커 헤드 좌표(±SNAP) · 선정 헤드 집합(K개) · 최원 유하거리 ·
    max_load · corridor 총연장

★두 계통은 **입력이 다르다**. A 는 DXF 엔티티에서 제 그래프를 만들고(자동 추출),
G 는 사람이 손질한 EditBoard 위에서 돈다. 그래서 숫자가 똑같기를 기대하는 것이
아니라, **다르면 그 차이가 설명되는지** 를 본다. 설명되지 않는 차이가 버그다.

    python tests/compare_module_a.py
"""
from __future__ import annotations

import json
import math
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_REPO = _ROOT.parent
for p in (str(_ROOT), str(_REPO)):
    if p not in sys.path:
        sys.path.insert(0, p)
os.chdir(_ROOT)

KEY = "B1F 현장조사 소화설비 평면도"
OUT = _ROOT / "tests" / "_out"
DOC = _REPO / "docs" / "module_g_vs_a.md"
SNAP_MM = 500.0        # 앵커 좌표 일치로 볼 여유


def g_side():
    """모듈 G — 손질 저장본 위에서 앵커 방식."""
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import select_and_expand

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not got.get("ok"):
        return {"ok": False, "error": got.get("error")}
    w = got["worst"]
    disks = es.board.disks
    an = w.get("anchor")
    return {
        "ok": True,
        "anchor_xy": (float(disks[an][0]), float(disks[an][1]))
        if an is not None and an < len(disks) else None,
        "heads_xy": [(float(disks[h][0]), float(disks[h][1]))
                     for h in w["heads"] if h < len(disks)],
        "far_m": w["far_m"], "max_load": w["max_load"],
        "total_m": w["total_m"], "span_m": w["span_m"],
        "excluded": got.get("excluded_heads", 0),
        "candidates": got.get("candidate_heads", 0),
        "total_heads": got.get("total_heads", 0),
    }


def a_side():
    """모듈 A — 제 DXF 파이프라인으로 자동 추출 후 최불리 30."""
    try:
        import remote30_prototype as A
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"모듈 A import 실패: {exc}"}

    spec_path = _ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec_path.read_text(encoding="utf-8")).get("source_dxf")
    if not src or not os.path.isfile(src):
        return {"ok": False, "error": f"원본 DXF 를 찾을 수 없습니다: {src}"}

    try:
        # A 의 실제 진입점 — 파일 내용 해시로 캐시하는 번들 파서.
        bundle = A.parse_dxf_bundle_cached(Path(src))
        ents = bundle.entities
        # 레이어 카테고리는 A 의 이름 사전이 정한다(라우트도 이 값을 넘긴다).
        layers = {ly.get("name"): A._categorize_layer(ly.get("name") or "")
                  for ly in (bundle.layers or [])}
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"A 파싱 실패: {type(exc).__name__} {exc}"}

    try:
        sel = A.select_worst30_heads(ents, layers, k=30)
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"A 선정 실패: {exc}"}

    heads = [h.pos for h in getattr(sel, "heads", [])]
    total = sum(L for *_e, L in getattr(sel, "edges", [])) / 1000.0
    span = 0.0
    if heads:
        xs = [p_[0] for p_ in heads]
        ys = [p_[1] for p_ in heads]
        span = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0
    return {"ok": True, "heads_xy": heads, "total_m": round(total, 2),
            "spread_m": round(span, 1), "edges": len(getattr(sel, "edges", [])),
            # A 의 비-anchored 경로는 «먼 순서» 라 앵커라는 개념이 없다.
            # 첫 헤드를 앵커로 부르면 거짓이 되므로 그렇게 쓰지 않는다.
            "anchor_xy": None,
            # [G13] 같은 선정 결과로 A 의 SDF 까지 뽑는다 — 두 번 고르지 않는다.
            "sel": sel}


def a_sdf(bundle_sel):
    """[G13] 같은 도면을 모듈 A 로 방출한 SDF.

    ★A 의 FX 실배관 materialize 는 끄고(pipe_entities=None) 돌린다. G 는 아직
      FX 를 실배관으로 펴지 않으므로, 켠 채로 견주면 「FX_<기하> Pipe-set 이
      더 있다」는 차이가 규격 바인딩 차이인 양 읽힌다.
    """
    if not bundle_sel:
        return None, "A 선정 결과 없음"
    try:
        import remote30_prototype as A
        tables = A.build_input_tables(bundle_sel, None,
                                      project_title="G13 대조 (모듈 A)")
        OUT.mkdir(parents=True, exist_ok=True)
        out = OUT / "g13_module_a.sdf"
        A.emit_sdf(tables, out, project_title="G13 대조 (모듈 A)")
        return out, None
    except Exception as exc:      # noqa: BLE001
        return None, f"{type(exc).__name__}: {exc}"


def sdf_shape(path):
    """SDF 의 «모양» — Pipe-set 구성·개수·좌표 폭. [G13] 대조용."""
    import xml.etree.ElementTree as ET
    r = ET.parse(str(path)).getroot()
    sets = [(ps.findtext("Pipe-type/Name"), len(ps.findall("Pipe")))
            for ps in r.iter("Pipe-set")]
    xs, ys = [], []
    for nd in r.iter("Node"):
        for q in nd.iter("Position"):
            try:
                xs.append(float(q.get("x")))
                ys.append(float(q.get("y")))
            except (TypeError, ValueError):
                pass
    lens = []
    for pp in r.iter("Pipe"):
        try:
            lens.append(float(pp.get("length") or 0.0))
        except ValueError:
            pass
    return {
        "pipe_sets": sets,
        "nodes": len(list(r.iter("Node"))),
        "pipes": len(list(r.iter("Pipe"))),
        "nozzles": len(list(r.iter("Nozzle"))),
        "span_x": (max(xs) - min(xs)) if xs else 0.0,
        "span_y": (max(ys) - min(ys)) if ys else 0.0,
        "length_sum": round(sum(lens), 3),
    }


def _fmt_sets(sh):
    return " / ".join(f"{n or '(빈칸)'}({c})" for n, c in sh["pipe_sets"])


def g13(a_selection):
    """[G13] 같은 도면의 두 SDF 를 견준다 — 구성·좌표 폭·위상 불변."""
    print("\n[G13] SDF 대조")
    from services.cad_import.design.bore import extract_dia_text_points
    from services.cad_import.design.emit import emit_design_sdf
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    from services.cad_import.edit.session import EditSession
    from services.cad_import.pipeline import handoff, stage1 as s1

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not got.get("ok"):
        print("  !! 제한 전개 실패:", got.get("error"))
        return None
    spec = _ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec.read_text(encoding="utf-8")).get("source_dxf")
    world = handoff.load_world(KEY, src, s1.World) if src else None
    texts = extract_dia_text_points(world.texts) if world else []
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], texts,
                              board_pts=es.board.pts)
    OUT.mkdir(parents=True, exist_ok=True)
    gs = sdf_shape(emit_design_sdf(tbl, OUT / "g13_module_g.sdf"))
    print(f"  G : 노드 {gs['nodes']} · 배관 {gs['pipes']} · 노즐 {gs['nozzles']}"
          f" · 길이합 {gs['length_sum']} m")
    print(f"      좌표 폭 x {gs['span_x']:.0f} · y {gs['span_y']:.0f}")
    print(f"      Pipe-set {len(gs['pipe_sets'])}개 — {_fmt_sets(gs)}")

    a_path, a_err = a_sdf(a_selection)
    rs = None
    if a_path is None:
        print(f"  A : 방출 못함 — {a_err}")
    else:
        rs = sdf_shape(a_path)
        print(f"  A : 노드 {rs['nodes']} · 배관 {rs['pipes']} · 노즐 {rs['nozzles']}"
              f" · 길이합 {rs['length_sum']} m")
        print(f"      좌표 폭 x {rs['span_x']:.0f} · y {rs['span_y']:.0f}")
        print(f"      Pipe-set {len(rs['pipe_sets'])}개 — {_fmt_sets(rs)}")

    verdict = {}
    if rs:
        g_names = [n for n, _c in gs["pipe_sets"]]
        r_names = [n for n, _c in rs["pipe_sets"]]
        # ★FX 슬롯은 둘의 «다른 일» 이다. A 는 헤드마다 신축배관을 실배관으로 펴서
        #   FX_<기하> Pipe-set 을 만들고, G 는 그 자리를 노즐 접속으로 둔 채 FX 를
        #   빈 정의로만 싣는다(지시서 G9-1 의 6종). 그러니 FX 로 시작하는 칸을
        #   빼고 견주는 것이 «라이브러리가 같은가» 라는 물음에 맞는 비교다.
        core = lambda ns: [n for n in ns if n and not n.startswith("FX")]
        g_core, r_core = core(g_names), core(r_names)
        ok_names = (g_core == r_core
                    and len(gs["pipe_sets"]) == len(rs["pipe_sets"])
                    and g_names[0] is None and r_names[0] is None)
        verdict["names"] = (ok_names, g_core, r_core)
        base = max(rs["span_x"], rs["span_y"]) or 1.0
        cur = max(gs["span_x"], gs["span_y"])
        ok_span = abs(cur - base) / base <= 0.05
        verdict["span"] = (ok_span, cur, base)
        _fx = lambda ns: [n for n in ns if n and n.startswith("FX")]
        print(f"  [{'OK  ' if ok_names else 'FAIL'}] 라이브러리 6종 동일 "
              f"(빈칸 + {g_core}) · 개수 {len(gs['pipe_sets'])}={len(rs['pipe_sets'])}")
        print(f"        FX 슬롯: G {_fx(g_names)} (빈 정의) vs "
              f"A {_fx(r_names)} (신축배관 실배관화)")
        print(f"  [{'OK  ' if ok_span else 'FAIL'}] 좌표 폭 ±5% "
              f"(G {cur:.0f} vs A {base:.0f})")

    # 위상 불변 — 종전 G 산출과 노드·배관·노즐·길이합이 같은가.
    # ★기준선은 G9 «이전» 방출기(커밋 9d581f8^)로 뽑은 값이다. 같은 세션에서 두 번
    #   돌려 같더라는 것은 재현성일 뿐 위상 불변의 증거가 아니다.
    bl = _ROOT / "tests" / "_g13_topology.json"
    keys = ("nodes", "pipes", "nozzles", "length_sum")
    cur_t = {k: gs[k] for k in keys}
    if bl.is_file():
        old_t = json.loads(bl.read_text(encoding="utf-8"))
        note = old_t.pop("_note", "")
        same = all(old_t.get(k) == cur_t[k] for k in keys)
        verdict["topology"] = (same, old_t, cur_t)
        print(f"  [{'OK  ' if same else 'FAIL'}] 위상 불변 (G9 이전 산출 대비) — "
              f"종전 {old_t} / 지금 {cur_t}")
    else:
        bl.write_text(json.dumps(cur_t, ensure_ascii=False), encoding="utf-8")
        verdict["topology"] = (None, None, cur_t)
        print(f"  [기록] 위상 기준선 신규 — {cur_t} (다음 실행부터 비교)")
    return {"g": gs, "a": rs, "a_err": a_err, "verdict": verdict}


# [G18] 보정 지시서 ② §2 가 적어 둔 수치. 이 보정은 표시/흐름만 건드렸으므로
#       아래가 하나라도 달라지면 위상이나 계산이 흔들린 것이다.
G18_REFERENCE = {
    "헤드": 30, "최원 유하거리 (m)": 171.87, "설계면적 폭 (m)": 54.25,
    "corridor 총연장 (m)": 207.31, "주배관 담당 헤드 수": 30,
    "노드": 62, "배관": 61, "노즐": 30, "기기": 0, "루프 잔여": 0,
    "관경 근거 — 도면 텍스트": 27, "관경 근거 — 별표1 보강": 0,
    "관경 근거 — 별표1 폴백": 34,
    "부속": 20, "부속 판정 불가": 3, "등가길이 미해결": 0,
}


def g18():
    """[G18] 처음부터 끝까지 한 번 돌린 값이 §2 기준선과 같은가."""
    print("")
    print("[G18] 보정 전 기준선과 대조")
    from services.cad_import.design.bore import extract_dia_text_points
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    from services.cad_import.edit.session import EditSession
    from services.cad_import.pipeline import handoff, stage1 as s1

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not got.get("ok"):
        print("  !! 제한 전개 실패:", got.get("error"))
        return None
    spec = _ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec.read_text(encoding="utf-8")).get("source_dxf")
    world = handoff.load_world(KEY, src, s1.World) if src else None
    texts = extract_dia_text_points(world.texts) if world else []
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], texts,
                              board_pts=es.board.pts,
                              excluded_heads=got.get("excluded_heads", 0))
    m = dict(tbl.meta)
    w = got["worst"]
    cur = {
        "헤드": len(w.get("heads") or []),
        "최원 유하거리 (m)": w.get("far_m"),
        "설계면적 폭 (m)": w.get("span_m"),
        "corridor 총연장 (m)": w.get("total_m"),
        "주배관 담당 헤드 수": w.get("max_load"),
        "노드": len(tbl.nodes), "배관": len(tbl.pipes),
        "노즐": len(tbl.nozzles), "기기": len(tbl.equipment),
        "루프 잔여": m.get("루프 잔여 배관(표 꼬리)"),
        "관경 근거 — 도면 텍스트": m.get("관경 근거 — 도면 텍스트"),
        "관경 근거 — 별표1 보강": m.get("관경 근거 — 별표1 보강 (text<min)"),
        "관경 근거 — 별표1 폴백": m.get("관경 근거 — 별표1 폴백 (text 없음)"),
        "부속": len(tbl.fittings),
        "부속 판정 불가": m.get("부속 판정 불가"),
        "등가길이 미해결": m.get("등가길이 미해결"),
    }
    diff = [(k, v, cur.get(k)) for k, v in G18_REFERENCE.items()
            if str(cur.get(k)) != str(v)]
    for k, v in G18_REFERENCE.items():
        mark = "OK  " if str(cur.get(k)) == str(v) else "FAIL"
        print(f"  [{mark}] {k} · 기준 {v} / 지금 {cur.get(k)}")
    print("  -> " + ("전부 같다" if not diff else f"다른 항목 {len(diff)}개"))
    return {"cur": cur, "ref": G18_REFERENCE, "same": not diff, "diff": diff}


def main() -> int:
    print("[G8] 모듈 A 대조\n")
    g = g_side()
    if not g.get("ok"):
        print("  !! G 쪽 실패:", g.get("error"))
        return 1
    print(f"  G : 앵커 {g['anchor_xy']} · 최원 {g['far_m']} m · "
          f"max_load {g['max_load']} · corridor {g['total_m']} m")
    print(f"      후보 {g['candidates']:,} / 도면 {g['total_heads']:,} "
          f"(전개가 못 붙여 제외 {g['excluded']:,})")

    a = a_side()
    rows = []
    if not a.get("ok"):
        print("  A : 돌리지 못함 —", a.get("error"))
        rows.append(("대조 상태", "G 단독", "A 미실행", a.get("error", "")))
    else:
        print(f"  A : 앵커 {a['anchor_xy']} · corridor {a['total_m']} m")
        # G 헤드의 퍼짐 — 설계면적인지 아닌지를 가르는 지표(D1).
        gxs = [p_[0] for p_ in g["heads_xy"]]
        gys = [p_[1] for p_ in g["heads_xy"]]
        g_spread = round(math.hypot(max(gxs) - min(gxs),
                                    max(gys) - min(gys)) / 1000.0, 1)
        rows.append(("선정 방식", "앵커 (D1)", "먼 순서 (A 기존)",
                     "지시서 D1 이 바꾸라고 한 지점"))
        rows.append(("선정 헤드 수", str(len(g["heads_xy"])),
                     str(len(a["heads_xy"])), ""))
        rows.append(("**헤드 퍼짐 (대각, m)**", str(g_spread),
                     str(a["spread_m"]),
                     "작을수록 한 설계구역 — G 가 뭉친다"))
        rows.append(("앵커 좌표", str(g["anchor_xy"]), "없음",
                     "A 비-anchored 경로엔 앵커 개념이 없다"))
        rows.append(("최원 유하거리 (m)", str(g["far_m"]), "—", ""))
        rows.append(("max_load", str(g["max_load"]), "—", ""))
        rows.append(("corridor 총연장 (m)", str(g["total_m"]),
                     str(a["total_m"]),
                     "A 는 흩어진 30개라 경로가 길다"))
        rows.append(("corridor 간선 수", "—", str(a["edges"]), ""))

    OUT.mkdir(parents=True, exist_ok=True)
    DOC.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 모듈 G vs 모듈 A — 최불리 배관망 대조", "",
        f"도면: `{KEY}`", "",
        "## 왜 숫자가 같기를 기대하지 않는가", "",
        "두 계통은 **입력이 다르다**. A 는 DXF 엔티티에서 제 그래프를 자동 추출하고,",
        "G 는 사람이 손질한 `EditBoard` 위에서 돈다(모듈 F 와 같은 망). 그래서 대조의",
        "목적은 «같은 값» 이 아니라 «다르면 그 차이가 설명되는가» 다.", "",
        "## G 결과", "",
        f"- 앵커 좌표 : `{g['anchor_xy']}`",
        f"- 최원 유하거리 : {g['far_m']} m",
        f"- 설계면적 폭 : {g['span_m']} m",
        f"- corridor 총연장 : {g['total_m']} m",
        f"- 주배관 담당 헤드 수(max_load) : {g['max_load']}",
        f"- 후보 헤드 : {g['candidates']:,} / 도면 {g['total_heads']:,}",
        f"- **전개가 붙이지 못해 제외한 헤드 : {g['excluded']:,}**", "",
        "## 대조", "",
        "| 항목 | G | A | 비고 |",
        "|---|---|---|---|",
    ]
    for r in rows:
        lines.append(f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} |")
    lines += [
        "", "## 설명되어야 하는 차이", "",
        "- **후보 집합** — G 는 전개가 배관에 붙일 수 있는 헤드에서만 고른다",
        "  (BLOCKED B4 1안). B1F 에서 619 / 3,163 이라 A 와 후보 모수가 다르다.",
        "- **망의 출처** — A 는 자동 추출망, G 는 손질망이다. 손질로 이어 붙인",
        "  배관이 있으면 G 쪽 도달 범위가 넓다.",
        "- **corridor 정의** — G 의 `total_m` 은 최단경로 합집합(board 간선)이고,",
        "  전개 배관 길이 합과는 구조적으로 다르다(BLOCKED B5).",
        "- **선정 방식 자체** — A 의 비-anchored 경로는 「급수원에서 먼 순서 K개」다.",
        "  지시서 D1 이 바꾸라고 한 지점이라 **결과가 다른 것이 정상**이다.",
        "  실측 퍼짐이 그 차이를 그대로 보여준다(A 134.1 m vs G 31.4 m).",
        "  A 에도 앵커 경로(`select_worst30_heads_anchored`)가 있으나 알람밸브",
        "  좌표와 헤드 영역이 둘 다 있어야 서고, 이 저장본에는 없어 세우지 못했다.", "",
        "관경이 다른 간선은 전부 `source`(text / nfpc_min / nfpc_fallback)로",
        "설명되어야 한다. 설명되지 않는 차이가 버그다.", "",
    ]
    _a_sel = a.get("sel") if a.get("ok") else None
    # [G13] SDF 구성 대조 — 규격 바인딩·좌표 정규화가 A 와 같은 모양인가.
    s13 = g13(_a_sel)
    if s13:
        gs, rs, vd = s13["g"], s13["a"], s13["verdict"]
        lines += ["", "## [G13] SDF 대조 — 같은 도면, 두 계통의 산출", ""]
        if rs is None:
            lines += [f"모듈 A 의 SDF 를 뽑지 못했다 — `{s13['a_err']}`. "
                      "G 쪽 값만 싣는다.", ""]
        lines += [
            "| 항목 | G | A |", "|---|---|---|",
            f"| Pipe-set 개수 | {len(gs['pipe_sets'])} | "
            f"{len(rs['pipe_sets']) if rs else '—'} |",
            f"| 첫 칸이 빈 placeholder | {gs['pipe_sets'][0][0] is None} | "
            f"{rs['pipe_sets'][0][0] is None if rs else '—'} |",
            f"| Pipe-type 구성 | {_fmt_sets(gs)} | {_fmt_sets(rs) if rs else '—'} |",
            f"| 노드 / 배관 / 노즐 | {gs['nodes']} / {gs['pipes']} / {gs['nozzles']} | "
            + (f"{rs['nodes']} / {rs['pipes']} / {rs['nozzles']} |" if rs else "— |"),
            f"| 좌표 폭 (x × y) | {gs['span_x']:.0f} × {gs['span_y']:.0f} | "
            + (f"{rs['span_x']:.0f} × {rs['span_y']:.0f} |" if rs else "— |"),
            f"| 배관 길이 합 (m) | {gs['length_sum']} | "
            + (f"{rs['length_sum']} |" if rs else "— |"), "",
            "### 수용 기준", "",
        ]
        if "names" in vd:
            ok, g_core, r_core = vd["names"]
            lines.append(f"- [{'PASS' if ok else 'FAIL'}] **Pipe-set 구성** — "
                         f"개수 {len(gs['pipe_sets'])} = "
                         f"{len(rs['pipe_sets'])}, 첫 칸 빈 placeholder 동일, "
                         f"라이브러리 6종의 이름·순서 동일: `{g_core}`.")
            lines.append("  마지막 한 칸만 다르다 — G 는 `FX`(빈 정의), A 는 "
                         "`FX_20A_216`(배관 30). A 는 헤드마다 신축배관을 **실배관으로 펴서** "
                         "규격 기하별 Pipe-set 을 동적 생성하고, G 는 그 자리를 노즐 접속으로 "
                         "둔 채 `FX` 를 드롭다운용 빈 정의로만 싣는다(지시서 G9-1 의 6종). "
                         "**계산의 차이가 아니라 신축배관을 펴느냐 마느냐의 차이**이며, "
                         "G 의 FX 실배관화는 이번 지시 범위 밖이다.")
        if "span" in vd:
            ok, cur, base = vd["span"]
            lines.append(f"- [{'PASS' if ok else 'FAIL'}] **좌표 폭 ±5%** — "
                         f"G {cur:.0f} · A {base:.0f} "
                         f"(차이 {abs(cur - base) / (base or 1) * 100:.1f}%). "
                         "둘 다 bbox 중심 → (0,0), 긴 축 → 캔버스 단위(3000) 규칙이다.")
        ok, old_t, cur_t = vd.get("topology", (None, None, None))
        if ok is None:
            lines.append(f"- [기록] **위상 불변** — 기준선을 새로 떴다 `{cur_t}`. "
                         "다음 실행부터 이 값과 비교한다.")
        else:
            lines.append(f"- [{'PASS' if ok else 'FAIL'}] **위상 불변** — "
                         f"G9 이전(`9d581f8^`) 방출기 산출 `{old_t}` / 지금 `{cur_t}`. "
                         "좌표 정규화와 아이소매트릭은 표시만 바꾼다는 증거다.")
        lines.append("")
    # [G18] 처음부터 끝까지 한 번 — §2 기준선과 견준다.
    r18 = g18()
    if r18:
        lines += ["", "## [G18] 보정 ② 전후 대조 (지시서 §2 기준선)", "",
                  "이 보정은 표시(SLF 경로·헤드 수직·좌표)와 흐름(산출물 선택)만",
                  "건드렸다. 아래가 하나라도 달라지면 위상이나 계산이 흔들린 것이다.",
                  "",
                  "| 항목 | 보정 전(§2) | 보정 후 | |", "|---|---|---|---|"]
        for k, v in r18["ref"].items():
            now = r18["cur"].get(k)
            same = "같음" if str(now) == str(v) else "**다름**"
            lines.append(f"| {k} | {v} | {now} | {same} |")
        lines += [
            "", "### 화면 증빙", "",
            "PIPENET 은 이 환경에서 띄울 수 없어 **그 화면 캡처는 내지 못했다.**",
            "대신 두 가지를 낸다.", "",
            "1. **미리보기 캡처** — 저장될 좌표를 그대로 그린 화면이다. 「미리보기",
            "   좌표 == 저장된 SDF 의 `<Position>`」은 `tests/test_design_dialog.py`",
            "   가 writer 자리수(`.6g`)로 62/62 문자열 일치까지 확인한다. 그래서 이",
            "   그림은 PIPENET 이 그릴 형태와 같은 좌표에서 나온 것이다.",
            "2. **파일 수준 확인** — PIPENET 화면에서 볼 `Type`·`Diameter` 두 열이",
            "   무엇을 보여줄지는 파일에서 직접 셀 수 있다(`tests/test_sdf_post.py`).",
            "   Type 이 빈 배관 0건 · 호칭경이 schedule 에 안 묶인 배관 0건 ·",
            "   User-lib 가 옆의 `.slf` 파일명 하나 · 그 SLF 가 쓰인 호칭경 6종",
            "   (25·40·50·65·100·150)을 전부 정의.", "",
            "![아이소매트릭 미리보기](images/module_g_preview_iso.png)", "",
            "*배관 굵기는 담당 헤드 수에 비례한다. 파란 원이 급수원, 붉은 △ 가 헤드다.*",
            "", "![배관표 미리보기](images/module_g_preview_table.png)", "",
            "*`관종` 이 전 행 `KSD 3507`, `호칭경` 이 65/25/150/40 로 채워진다 —",
            "PIPENET 의 `Type`·`Diameter` 열이 보여 줄 값이 이것이다. `관경 근거` 는",
            "이 보정에서 이어 붙인 열이다.*", "",
            "캡처는 `python tests/_capture_preview.py` 로 다시 만든다.", "",
        ]
        tail = ("전부 같다" if r18["same"]
                else "다른 항목 " + str(len(r18["diff"])) + "개")
        lines += ["", f"-> **{tail}**", ""]
    DOC.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n  대조표: {DOC}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
