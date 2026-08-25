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


# ─────────────────────────────────────────────────────────── G10
def g10():
    print("\n[G10] 관종 선택을 표까지 잇기")
    from services.cad_import.design.sdf_post import UnknownSchedule
    from services.cad_import.design.tables import build_design_tables

    net = {"pipe_data": {"P1": {"start": "N1", "end": "N2", "length_m": 1.0}},
           "nodes_meta_runtime": {
               "N1": {"coords": [0.0, 0.0, 0.0], "type_id": "pump"},
               "N2": {"coords": [1.0, 0.0, 0.0], "type_id": "base"}}}
    worst = {"heads": [], "loads": {(0, 1): 1}}

    t = build_design_tables(net, worst, {"P1": (0, 1)}, [])
    check("기본 관종은 KSD 3507", t.pipes[0]["type"] == "KSD 3507",
          t.pipes[0]["type"])
    check("meta 에 관종이 남는다", ("배관 규격(기본)", "KSD 3507") in t.meta)

    t2 = build_design_tables(net, worst, {"P1": (0, 1)}, [],
                             default_schedule="CPVC2")
    check("기본값을 바꾸면 표가 따라온다", t2.pipes[0]["type"] == "CPVC2",
          t2.pipes[0]["type"])

    t3 = build_design_tables(net, worst, {"P1": (0, 1)}, [],
                             schedule_by_pipe={"P1": "KSD 3576"})
    check("배관별 지정이 먹는다", t3.pipes[0]["type"] == "KSD 3576",
          t3.pipes[0]["type"])

    check("없는 이름은 오류(조용히 기본값 아님)",
          _raises(lambda: build_design_tables(net, worst, {"P1": (0, 1)}, [],
                                              default_schedule="KSD3507"),
                  UnknownSchedule), "공백 빠진 이름")

    # 없는 이름을 주면 **파일을 만들지 않아야** 한다.
    from services.cad_import.design.emit import emit_design_sdf
    bad = OUT / "should_not_exist_g10.sdf"
    if bad.exists():
        bad.unlink()
    t_bad = build_design_tables(net, worst, {"P1": (0, 1)}, [])
    t_bad.pipes[0]["type"] = "NOPE"
    check("잘못된 관종이면 파일을 안 만든다",
          _raises(lambda: emit_design_sdf(t_bad, bad), UnknownSchedule)
          and not bad.exists(), str(bad.name))
    return True



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
# ─────────────────────────────────────────────────────────── G14
_SYNTH_SDF = r'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE Project SYSTEM "spray.dtd">
<Project version="1.8  (0)">
  <Network-spray>
    <Title>우리 제목</Title>
    <Title>남의 제목 A</Title>
    <Title>남의 제목 B</Title>
    <Network-description>남의 설명</Network-description>
    <Libraries>
      <User-lib file="C:\Users\someone\Desktop\alpha.slf"/>
      <User-lib file="D:\lib\beta.slf"/>
    </Libraries>
    <Nodes/>
    <Links/>
  </Network-spray>
  <Graphics>
    <Text-element><Text>주기 1</Text></Text-element>
    <Text-element><Text>주기 2</Text></Text-element>
    <Text-element><Text>주기 3</Text></Text-element>
    <Display-options/>
  </Graphics>
</Project>
'''


def g14():
    print("\n[G14] 템플릿 잔재 정리와 SLF 경로 재작성")
    from services.cad_import.design.sdf_post import sanitize_template

    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / "g14_synth.sdf"
    p.write_text(_SYNTH_SDF, encoding="utf-8")
    got = sanitize_template(p, "g14_synth.slf")

    r = ET.parse(str(p)).getroot()
    libs = [ul.get("file") for lb in r.iter("Libraries")
            for ul in lb.findall("User-lib")]
    check("User-lib 하나만 남고 동봉 SLF 를 가리킨다",
          libs == ["g14_synth.slf"], str(libs))
    check("절대경로가 아니라 파일명뿐",
          libs and "\\" not in libs[0] and "/" not in libs[0], str(libs))
    te = [1 for g in r.iter("Graphics") for _ in g.findall("Text-element")]
    check("Text-element 0", len(te) == 0, f"{len(te)}개 남음")
    titles = [t.text for ns in r.iter("Network-spray") for t in ns.findall("Title")]
    check("Title 은 첫 개만", titles == ["우리 제목"], str(titles))
    nd = [1 for ns in r.iter("Network-spray") for _ in ns.findall("Network-description")]
    check("Network-description 0", len(nd) == 0, f"{len(nd)}개 남음")
    check("지운 수를 정직하게 돌려준다",
          got == {"user_lib": 2, "text_element": 3, "title": 2, "net_desc": 1},
          str(got))
    head = p.read_text(encoding="utf-8").splitlines()[:2]
    check("DOCTYPE 보존", len(head) > 1
          and 'DOCTYPE Project SYSTEM "spray.dtd"' in head[1], " / ".join(head))
    return True


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

    # ★호칭경이 제 schedule 의 Pipe-size 에 실제로 있는가. 없으면 PIPENET 이
    #   내경을 못 찾아 Diameter 가 "Unset" 이 된다 — G9 수용 기준인데 지금까지
    #   눈으로만 봤다. 관경 판정이 0 을 흘려도 여기서 잡힌다.
    unbound, no_type = [], 0
    for links in root.iter("Links"):
        for ps in links.findall("Pipe-set"):
            nm = ps.findtext("Pipe-type/Name")
            sizes = {round(float(e.get("size")), 6)
                     for e in ps.findall("Pipe-type/Pipe-size")}
            for pipe in ps.findall("Pipe"):
                if not nm:
                    no_type += 1
                    continue
                b = round(float(pipe.get("bore") or 0), 6)
                if b not in sizes:
                    unbound.append((pipe.get("label"), b, nm))
    check("Type 열이 빈 배관이 없다('None defined' 자리)", no_type == 0,
          f"{no_type}개")
    check("호칭경이 schedule 에 묶인다('Unset' 자리)", not unbound,
          f"안 묶인 것 {unbound[:3]}" if unbound else "61개 전부")

    # ★DOCTYPE — 없으면 «일부» PIPENET 설치에서만 안 열린다. 내 화면에서 열렸다는
    #   것이 증거가 못 된다. 모듈 A 가 같은 이유로 헤더를 직접 붙인다.
    head = plain.read_text(encoding="utf-8")[:120].splitlines()[:2]
    check("SDF 머리에 DOCTYPE 이 있다",
          len(head) > 1 and 'DOCTYPE Project SYSTEM "spray.dtd"' in head[1],
          " / ".join(head))
    check("XML 선언이 레퍼런스와 같은 표기",
          head and head[0] == '<?xml version="1.0" encoding="UTF-8"?>',
          head[0] if head else "(없음)")

    # ★[G14] 산출 SDF 가 남의 라이브러리를 가리키면 Type 은 채워져도 Diameter 가
    #   전부 "Unset" 이 된다 — 관종 바인딩은 됐는데 관경만 안 뜨는 모습이라
    #   원인을 엉뚱한 데서 찾게 된다.
    libs = [ul.get("file") for lb in root.iter("Libraries")
            for ul in lb.findall("User-lib")]
    check("User-lib 는 동봉 SLF 하나뿐", libs == [plain.with_suffix(".slf").name],
          str(libs))
    check("경로가 아니라 파일명 — 폴더를 옮겨도 산다",
          bool(libs) and "\\" not in libs[0] and "/" not in libs[0], str(libs))
    check("가리킨 SLF 가 옆에 실제로 있다",
          (plain.parent / libs[0]).is_file() if libs else False,
          libs[0] if libs else "(없음)")

    txt = plain.read_text(encoding="utf-8", errors="replace")
    leaks = {w: txt.count(w) for w in
             ("Officetell", "WATER TANK", "3-1 type", "PH1F")}
    check("남의 프로젝트 정보가 안 남는다", not any(leaks.values()), str(leaks))
    titles = [t.text for ns in root.iter("Network-spray") for t in ns.findall("Title")]
    check("제목은 우리 것 하나뿐", len(titles) == 1, str(titles))

    # 폴더째 옮겨도 관경이 산다 — 상대 참조라야 성립한다.
    import shutil as _sh
    moved = OUT / "g14_moved"
    if moved.is_dir():
        _sh.rmtree(moved)
    moved.mkdir(parents=True)
    _sh.copyfile(plain, moved / plain.name)
    _sh.copyfile(plain.with_suffix(".slf"), moved / plain.with_suffix(".slf").name)
    mr = ET.parse(str(moved / plain.name)).getroot()
    mlibs = [ul.get("file") for lb in mr.iter("Libraries")
             for ul in lb.findall("User-lib")]
    check("다른 폴더로 옮겨도 라이브러리를 찾는다",
          bool(mlibs) and (moved / mlibs[0]).is_file(),
          f"{mlibs} · 옆에 있음 {bool(mlibs) and (moved / mlibs[0]).is_file()}")

    # ★같은 표로 두 번 저장 — 창은 표를 들고 있다가 다시 저장할 수 있다.
    #   방출이 표를 in-place 로 굽으면 두 번째 저장이 어긋난다.
    t = tables()
    once = emit_design_sdf(t, OUT / "reg_twice_a.sdf", iso=True, iso_z_scale=2.0)
    twice = emit_design_sdf(t, OUT / "reg_twice_b.sdf", iso=True, iso_z_scale=2.0)

    def pos(p):
        r = ET.parse(p).getroot()
        return [(n.get("label"), q.get("x"), q.get("y"))
                for n in r.iter("Node") for q in n.iter("Position")]

    check("같은 표로 두 번 저장해도 그림이 같다", pos(once) == pos(twice),
          "두 번째가 어긋남" if pos(once) != pos(twice) else "동일")
    # 껐다 저장하면 평면으로 돌아와야 한다 — 첫 저장의 등각이 남으면 안 된다.
    off = emit_design_sdf(t, OUT / "reg_twice_c.sdf")
    check("아이소 저장 뒤 평면 저장이 평면이다", pos(off) == pos(plain),
          "등각이 남았다" if pos(off) != pos(plain) else "동일")
    return True


def main() -> int:
    for fn in (g9, g10, g11, g12, g14, regression):
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
