# -*- coding: utf-8 -*-
"""[G9~G12] SDF 후처리 검증 — 규격 바인딩·좌표 정규화·아이소매트릭.

★가장 중요한 것은 **회귀**다. 이 보정이 건드리는 것은 표시 좌표와 Pipe-type 뿐이고,
수리계산 입력(length · rise · elevation · bore · 부속 · 노즐 유량)은 **한 개도**
바뀌면 안 된다. 그것을 diff 로 증명한다.

    python tests/test_sdf_post.py
"""
from __future__ import annotations

import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)

KEY = "B1F 현장조사 소화설비 평면도"
OUT = _ROOT / "tests" / "_out"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


class _T:
    """합성 테이블 — PipeTablesG 와 같은 모양만 있으면 된다."""

    def __init__(self, nodes, pipes=None):
        self.nodes = nodes
        self.pipes = pipes or []
        self.nozzles = []
        self.fittings = []
        self.equipment = []
        self.meta = []


# ─────────────────────────────────────────────────────────── G9
def g9():
    print("\n[G9] 배관 규격(Pipe-type) 주입")
    from services.cad_import.design.sdf_post import (
        SCHEDULE_DEFS, SCHEDULE_NAMES, UnknownSchedule, check_schedule,
        inject_pipe_types)

    check("6종 스케줄 정의", len(SCHEDULE_DEFS) == 6, " · ".join(SCHEDULE_NAMES))
    # 이름은 SLF 의 Item-name 과 철자·공백까지 같아야 한다.
    slf = _ROOT.parent / "assets" / "2. Pipenet_hand_FX28.slf"
    if slf.is_file():
        txt = slf.read_text(encoding="utf-8", errors="replace")
        miss = [n for n in SCHEDULE_NAMES if f">{n}<" not in txt]
        check("모든 스케줄 이름이 SLF 에 있다", not miss, f"없음 {miss}")

    check("모르는 이름은 오류", _raises(lambda: check_schedule("KSD3507"),
                                       UnknownSchedule), "공백 빠진 이름")
    check("정확한 이름은 통과", check_schedule("CPVC2") == "CPVC2")

    # ── 합성: 두 규격이 섞이면 Pipe-set 이 갈린다
    from services.cad_import.design.emit import _pc_models
    m, write_sdf = _pc_models()
    net = m.PipeNetwork(title="G9 합성")
    for i in range(4):
        net.nodes[str(i)] = m.Node(node_id=str(i), x=float(i) * 10,
                                   y=0.0, z=0.0, node_type="base")
    for i, sched in enumerate(["KSD 3507", "KSD 3507", "CPVC2"]):
        net.pipes[f"P{i}"] = m.Pipe(pipe_id=f"P{i}", from_node=str(i),
                                    to_node=str(i + 1), diameter_m=0.05,
                                    length_m=1.0, rise_m=0.0, material=sched)
    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / "g9_synth.sdf"
    from services.cad_import.design.emit import resolve_template_sdf
    write_sdf(net, p, template_path=resolve_template_sdf())
    inject_pipe_types(p, {"P0": "KSD 3507", "P1": "KSD 3507", "P2": "CPVC2"})

    root = ET.parse(p).getroot()
    sets = root.findall(".//Links/Pipe-set")
    named = [(s.findtext("Pipe-type/Name"), len(s.findall("Pipe"))) for s in sets]
    check("맨 앞은 빈 placeholder", named[0] == (None, 0), str(named[0]))
    used = {n: c for n, c in named if c}
    check("규격별로 Pipe-set 이 갈린다",
          used == {"KSD 3507": 2, "CPVC2": 1}, str(used))
    check("안 쓰인 규격도 정의만 노출(드롭다운)",
          len([n for n, c in named if n and not c]) == 4,
          f"빈 정의 {[n for n, c in named if n and not c]}")
    return True


def _raises(fn, exc):
    try:
        fn()
    except exc:
        return True
    except Exception:
        return False
    return False


# ─────────────────────────────────────────────────────────── G11
def g11():
    print("\n[G11] 좌표 정규화")
    from services.cad_import.design.sdf_post import normalize_node_coords

    t = _T([{"label": "1", "x": 100000.0, "y": 200000.0, "elevation": 0.0,
             "display_z": 5000.0},
            {"label": "2", "x": 700000.0, "y": 300000.0, "elevation": 3.0}])
    scale = normalize_node_coords(t, canvas_units=3000.0)
    xs = [n["x"] for n in t.nodes]
    ys = [n["y"] for n in t.nodes]
    check("가장 긴 축이 캔버스 단위", abs((max(xs) - min(xs)) - 3000.0) < 1e-6,
          f"x폭 {max(xs)-min(xs):.1f}")
    check("중심이 (0,0)", abs(sum(xs) / 2) < 1e-6 and abs(sum(ys) / 2) < 1e-6,
          f"중심 ({sum(xs)/2:.3f}, {sum(ys)/2:.3f})")
    check("display_z 가 같은 배율", abs(t.nodes[0]["display_z"] - 5000.0 * scale) < 1e-6,
          f"{t.nodes[0]['display_z']:.3f} = 5000×{scale:.6f}")
    check("elevation 은 안 건드린다",
          t.nodes[0]["elevation"] == 0.0 and t.nodes[1]["elevation"] == 3.0)
    return True


# ─────────────────────────────────────────────────────────── G12
def g12():
    print("\n[G12] 아이소매트릭 베이크")
    from services.cad_import.design.sdf_post import bake_isometric

    # 표고가 모두 같은 평면망 — lift 가 0 이어야 하고 예외 없이 돌아야 한다.
    flat = _T([{"label": "1", "x": 0.0, "y": 0.0, "elevation": 0.0},
               {"label": "2", "x": 100.0, "y": 0.0, "elevation": 0.0}])
    bake_isometric(flat)
    check("평면 전용 망에서 lift=0",
          abs(flat.nodes[1]["y"] - 100.0 * 0.5) < 1e-9,
          f"y={flat.nodes[1]['y']:.4f} (= x·sin30)")

    # 표고가 있으면 세로로 벌어지고, 배율을 올리면 그 폭만 커진다.
    def mk():
        return _T([{"label": "1", "x": 0.0, "y": 0.0, "elevation": 0.0},
                   {"label": "2", "x": 100.0, "y": 0.0, "elevation": 10.0}])

    a, b = mk(), mk()
    bake_isometric(a, iso_z_scale=1.0)
    bake_isometric(b, iso_z_scale=2.0)
    check("iso_z_scale 이 세로 분리만 키운다",
          abs(a.nodes[1]["x"] - b.nodes[1]["x"]) < 1e-9
          and abs(b.nodes[1]["y"] - a.nodes[1]["y"]) > 1.0,
          f"x 같음 {a.nodes[1]['x']:.2f} · y {a.nodes[1]['y']:.2f} → {b.nodes[1]['y']:.2f}")

    # ref_label 을 주면 그 노드에서 lift 가 정확히 0 이 된다(이음매 찢어짐 방지).
    c = mk()
    bake_isometric(c, ref_label="2")
    check("ref_label 노드에서 lift=0",
          abs(c.nodes[1]["y"] - (100.0 * 0.5)) < 1e-9,
          f"y={c.nodes[1]['y']:.4f}")

    # no_lift 노드는 표고 lift 를 건너뛴다.
    d = mk()
    bake_isometric(d, no_lift_labels={"2"})
    check("no_lift 노드는 lift 건너뜀",
          abs(d.nodes[1]["y"] - (100.0 * 0.5)) < 1e-9,
          f"y={d.nodes[1]['y']:.4f}")
    check("elevation 은 안 건드린다", d.nodes[1]["elevation"] == 10.0)
    return True


# ────────────────────────────────────────────── 회귀 (가장 중요)
def _calc_view(path):
    """SDF 에서 «수리계산에 쓰이는 값» 만 뽑는다. 좌표·Pipe-type 은 뺀다."""
    root = ET.parse(path).getroot()
    pipes = {}
    for p in root.iter("Pipe"):
        pipes[p.get("label")] = {
            k: p.get(k) for k in ("bore", "length", "rise", "roughness-or-c",
                                  "input", "output", "status")}
        pipes[p.get("label")]["fittings"] = sorted(
            (f.get("type"), f.get("count")) for f in p.iter("Fitting"))
    nodes = {n.get("label"): n.get("elevation") for n in root.iter("Node")}
    noz = {n.get("label"): (n.get("input"), n.get("output"),
                            (n.find("Flow-define").get("flow")
                             if n.find("Flow-define") is not None else None))
           for n in root.iter("Nozzle")}
    return {"pipes": pipes, "nodes": nodes, "nozzles": noz}


def regression():
    print("\n[회귀] 표시만 바뀌고 계산값은 그대로인가 — 실도면 B1F")
    from services.cad_import.design.bore import extract_dia_text_points
    from services.cad_import.design.emit import emit_design_sdf
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    from services.cad_import.edit.session import EditSession
    from services.cad_import.pipeline import handoff, stage1 as s1
    import json

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not check("제한 전개", got.get("ok"), str(got.get("error"))[:70]):
        return False
    spec = _ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec.read_text(encoding="utf-8")).get("source_dxf")
    world = handoff.load_world(KEY, src, s1.World)
    texts = extract_dia_text_points(world.texts) if world else []

    def tables():
        return build_design_tables(got["kfp"], got["worst"], got["edge_ref"],
                                   texts, board_pts=es.board.pts)

    OUT.mkdir(parents=True, exist_ok=True)
    plain = emit_design_sdf(tables(), OUT / "reg_plain.sdf")
    iso = emit_design_sdf(tables(), OUT / "reg_iso.sdf", iso=True,
                          iso_z_scale=2.0)
    big = emit_design_sdf(tables(), OUT / "reg_big.sdf", canvas_units=8000.0)

    a, b, c = _calc_view(plain), _calc_view(iso), _calc_view(big)
    check("아이소매트릭 켬/끔 — 계산값 동일", a == b,
          "다름" if a != b else f"배관 {len(a['pipes'])} · 노즐 {len(a['nozzles'])}")
    check("캔버스 크기 바꿔도 계산값 동일", a == c,
          "다름" if a != c else "동일")

    def span(p):
        r = ET.parse(p).getroot()
        xs = [float(q.get("x")) for n in r.iter("Node") for q in n.iter("Position")]
        return max(xs) - min(xs)

    check("좌표는 실제로 달라진다(표시가 바뀐 증거)",
          abs(span(plain) - span(big)) > 1.0,
          f"기본 {span(plain):.0f} · 캔버스8000 {span(big):.0f}")
    root = ET.parse(plain).getroot()
    ios = sorted({n.get("io-node") for n in root.iter("Node")})
    check("io-node 가 규약값뿐", set(ios) <= {"Input", "No"}, str(ios))
    names = [s.findtext("Pipe-type/Name") for s in root.findall(".//Links/Pipe-set")]
    check("Pipe-type 6종 노출 + placeholder",
          names[0] is None and len([n for n in names if n]) == 6, str(names))
    return True


def main() -> int:
    for fn in (g9, g11, g12, regression):
        fn()
    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("SDF 후처리 수용 기준 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
