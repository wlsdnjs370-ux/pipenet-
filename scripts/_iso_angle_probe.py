# -*- coding: utf-8 -*-
"""[아이소 각도] 등각 격자를 벗어난 배관을 찾는다 — 지시서 §7.

`ModuleF_헤드배관_꼬임_수정지시서.md` 의 계측 도구. 조치 전후로 **같은 표**를
낸다 — 단계 간 비교가 이 작업의 근거다.

판정의 뿌리:

    x' = (x − y)·cos30      y' = (x + y)·sin30 + (e − e_ref)·lift

이므로 평면 가로 → 화면 +30° · 평면 세로 → 화면 +150° · 헤드 스텁 → 수직.
**그 셋 말고 다른 각도는 나올 수 없다.** 나오면 평면 형상이 비직각이라는 뜻이다.

    python scripts/_iso_angle_probe.py [SDF경로]
"""
from __future__ import annotations

import math
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
DEF = (ROOT / "data" / "uploads" / "module_f" / "86e6fef47b2542e8_design"
       / "1. 입력도면 대명동 단위세대 평면도_수리계산입력.sdf")

COS30, SIN30 = 0.8660254037844387, 0.5


def _tag(e):
    return e.tag.split("}")[-1]


def read_sdf(path: Path):
    """Node/Pipe/Nozzle — 필요한 것만. 좌표는 Position, 표고는 attribute."""
    root = ET.parse(str(path)).getroot()
    nodes, pipes, heads = {}, [], set()
    for e in root.iter():
        t = _tag(e)
        if t == "Node":
            lab = e.attrib.get("label")
            pos = e.find("Position")
            if lab and pos is not None:
                nodes[str(lab)] = {
                    "x": float(pos.attrib.get("x", 0)),
                    "y": float(pos.attrib.get("y", 0)),
                    "e": float(e.attrib.get("elevation", 0) or 0),
                    "io": e.attrib.get("io-node", "No"),
                }
        elif t == "Pipe":
            a, b = e.attrib.get("input"), e.attrib.get("output")
            if a and b and "@" not in a and "@" not in b:
                pipes.append({"label": e.attrib.get("label", ""),
                              "a": str(a), "b": str(b),
                              "len": float(e.attrib.get("length", 0) or 0)})
        elif t == "Nozzle":
            a = e.attrib.get("input")
            if a:
                heads.add(str(a))
    return nodes, pipes, heads


def _ang(dx, dy):
    """화면 각도(0~180). 수평=0 · 수직=90."""
    return math.degrees(math.atan2(dy, dx)) % 180.0


def _cls(a, tol=1.5):
    if abs(a - 30.0) <= tol:
        return "+30° (평면 가로)"
    if abs(a - 150.0) <= tol:
        return "+150° (평면 세로)"
    if abs(a - 90.0) <= tol:
        return "수직 (헤드 스텁)"
    if a <= tol or a >= 180.0 - tol:
        return "수평 ← 있으면 안 됨"
    return "기타"


def _unbake(x, y):
    """등각 → 평면 되돌리기(표고 0 전제 · lift 항 없음).

        x' = (x−y)·cos30 · y' = (x+y)·sin30
        →  x = x'/(2cos30) + y'/(2sin30) · y = -x'/(2cos30) + y'/(2sin30)
    """
    u = x / (2 * COS30)
    v = y / (2 * SIN30)
    return u + v, v - u


def _seg_cross(p1, p2, p3, p4):
    def d(o, a, b):
        return ((a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]))
    d1, d2 = d(p3, p4, p1), d(p3, p4, p2)
    d3, d4 = d(p1, p2, p3), d(p1, p2, p4)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def main() -> int:
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    if not path.is_file():
        print(f"표본 없음: {path}")
        return 0
    nodes, pipes, heads = read_sdf(path)
    print(f"■ {path.name}")
    print(f"  절점 {len(nodes)} · 배관 {len(pipes)} · 노즐 {len(heads)}")

    # ── ① 화면 각도 분류
    rows = []
    for p in pipes:
        a, b = nodes.get(p["a"]), nodes.get(p["b"])
        if not a or not b:
            continue
        dx, dy = b["x"] - a["x"], b["y"] - a["y"]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            continue
        rows.append((p, a, b, _ang(dx, dy), L))
    cnt = Counter(_cls(r[3]) for r in rows)
    print("  [화면 각도]")
    for k in ("+30° (평면 가로)", "+150° (평면 세로)", "수직 (헤드 스텁)",
              "수평 ← 있으면 안 됨", "기타"):
        if cnt.get(k):
            print(f"    {k:22} {cnt[k]}")
    off = [r for r in rows if _cls(r[3]) in ("수평 ← 있으면 안 됨", "기타")]
    print(f"    ★등각 격자 벗어남 {len(off)}")
    if off:
        etc = Counter(round(r[3], 1) for r in off if _cls(r[3]) == "기타")
        if etc:
            print(f"       기타 각도: {sorted(etc.items())[:12]}")

    # ── ② 평면 복원 — 비직각 · 챔퍼 판정
    plan = {lab: _unbake(n["x"], n["y"]) for lab, n in nodes.items()}
    # 화면 1 m — 표 length 와 좌표 거리의 비(수평에 가까운 배관 중앙값)
    per_m = []
    for p, a, b, ang, L in rows:
        if p["len"] > 0.05 and abs(a["e"] - b["e"]) < 1e-9:
            per_m.append(L / p["len"])
    per_m.sort()
    unit = per_m[len(per_m) // 2] if per_m else 1.0
    print(f"  화면 1 m = {unit:.1f} 단위 (평면 기준 · 표본 {len(per_m)})")

    nb: dict[str, list] = {}
    for p in pipes:
        nb.setdefault(p["a"], []).append(p)
        nb.setdefault(p["b"], []).append(p)

    def plan_ang(p):
        (x1, y1), (x2, y2) = plan[p["a"]], plan[p["b"]]
        return _ang(x2 - x1, y2 - y1), math.hypot(x2 - x1, y2 - y1)

    # ★헤드 접속관은 평면 판정에서 **뺀다.**
    #
    #   평면에서 헤드와 부모는 같은 점이고 접속관 길이는 0 이다. 그런데 등각에서
    #   헤드만 수직으로 올라가므로, 되돌리기(lift 항을 모르는 역변환)는 그 수직선을
    #   평면 +45° 대각 424 mm 로 «복원» 한다 — 없던 비직각 30개다.
    #   처음에 이걸 안 빼서 비직각 75(=44+30+1) · X 교차 8(실제 2)로 부풀었다.
    flat = []
    for p in pipes:
        if p["a"] not in plan or p["b"] not in plan:
            continue
        if p["a"] in heads or p["b"] in heads:
            continue
        pa, L = plan_ang(p)
        if L < 1.0:
            continue
        flat.append((p, pa, L))
    nonrect = [(p, pa, L) for p, pa, L in flat
               if not (pa <= 1.5 or abs(pa - 90) <= 1.5 or pa >= 178.5)]
    print(f"  [평면] 길이 있는 배관 {len(flat)} · **비직각 {len(nonrect)}**")

    # 헤드로부터 홉수 — 신축배관 판정의 핵심 지표
    depth: dict[str, int] = {h: 0 for h in heads}
    frontier = list(heads)
    while frontier:
        nxt = []
        for lab in frontier:
            for p in nb.get(lab, ()):
                other = p["b"] if p["a"] == lab else p["a"]
                if other not in depth:
                    depth[other] = depth[lab] + 1
                    nxt.append(other)
        frontier = nxt
    hop = Counter()
    lens = Counter()
    for p, pa, L in nonrect:
        h = min(depth.get(p["a"], 99), depth.get(p["b"], 99))
        hop[h] += 1
        lens[round(L / unit * 1000)] += 1     # 평면 mm
    print(f"    헤드로부터 홉수: {dict(sorted(hop.items())[:6])}")
    print(f"    길이(mm) 상위: {lens.most_common(6)}")

    # ── ③ 평면 X 교차 (헤드 접속관 제외)
    segs = [(p["a"], p["b"], plan[p["a"]], plan[p["b"]]) for p, _pa, _L in flat]
    cross = []
    for i in range(len(segs)):
        for j in range(i + 1, len(segs)):
            a1, b1, p1, p2 = segs[i]
            a2, b2, p3, p4 = segs[j]
            if {a1, b1} & {a2, b2}:
                continue
            if _seg_cross(p1, p2, p3, p4):
                cross.append((f"{a1}-{b1}", f"{a2}-{b2}"))
    print(f"  [평면] ★X 교차 {len(cross)}건 {cross[:4]}")

    # ── ④ 헤드가 세로로 섰는가 · 접속관 겹침
    bad_stand = []
    stub_pairs = []
    for h in sorted(heads):
        ps = nb.get(h, [])
        if len(ps) != 1:
            continue
        par = ps[0]["b"] if ps[0]["a"] == h else ps[0]["a"]
        if h in nodes and par in nodes:
            if abs(nodes[h]["x"] - nodes[par]["x"]) > 1e-6:
                bad_stand.append(h)
            stub_pairs.append((h, par))
    print(f"  헤드 {len(heads)} · **세로로 안 선 헤드 {len(bad_stand)}**"
          f"{' ' + str(bad_stand[:6]) if bad_stand else ''}")

    # ── ⑤ 화면 교차 — 스텁이 낀 것과 배관끼리를 **갈라서** 센다(지시서 §4-⑴).
    #    스텁이 낀 것은 등각의 원리적 성질이라 고칠 대상이 아니고,
    #    배관끼리는 평면에 원래 있던 것(A 항목)이라 성격이 다르다.
    stub_set = {h for h, _p in stub_pairs}
    scr = []
    for p, a, b, _ang_, _L in rows:
        scr.append((p["a"], p["b"], (a["x"], a["y"]), (b["x"], b["y"]),
                    p["a"] in stub_set or p["b"] in stub_set))
    sv = pv = 0
    items = []
    for i in range(len(scr)):
        for j in range(i + 1, len(scr)):
            a1, b1, p1, p2, s1 = scr[i]
            a2, b2, p3, p4, s2 = scr[j]
            if {a1, b1} & {a2, b2}:
                continue
            if not _seg_cross(p1, p2, p3, p4):
                continue
            if s1 or s2:
                sv += 1
            else:
                pv += 1
            if len(items) < 8:
                items.append(f"{a1}-{b1} × {a2}-{b2}")
    print(f"  [등각] 화면 교차 {sv + pv}건 — 스텁이 낀 것 {sv}"
          f" · 배관끼리 {pv} {items[:4]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
