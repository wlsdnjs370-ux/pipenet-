# -*- coding: utf-8 -*-
"""D3 검증 — 만든 PDF 에서 글자를 도로 뽑아 결과 XML 원문과 1:1 대조한다.

렌더러 코드를 한 줄도 쓰지 않는다. SDF <Units> 와 XDSET <TABLE> 을 직접 파싱하고
환산 계수는 여기서 다시 적어 곱한다. 렌더러가 값을 훼손하면 여기서 어긋난다.
"""
from __future__ import annotations

import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from pypdf import PdfReader  # noqa: E402

SUB = _ROOT / "routes" / "제출용[최종]"
OUT = _ROOT / "data" / "_d3"

# 정의값. 렌더러의 UNIT_FACTORS 를 참조하지 않고 여기서 다시 적는다.
KGF_CM2 = 98066.5      # Pa
L_MIN = 60000.0        # m3/s → l/min
MM = 1000.0            # m → mm
ATM = 101325.0         # XDSET 노드 압력은 절대압, 관로 결과는 게이지압이다


def table(root, name):
    for t in root.iter("TABLE"):
        if t.get("name") == name:
            fields = [f.get("name") for f in t.iter("FIELD")]
            rows = [[(td.text or "").strip() for td in tr] for tr in t.iter("TR")]
            return fields, [dict(zip(fields, r)) for r in rows]
    return [], []


def key(label):
    """XDSET 은 같은 관로를 'M1/0' 으로도 적는다. 앞 토막이 SDF 라벨이다."""
    return label.split("/")[0].strip().upper()


def num(s):
    try:
        v = float(s)
    except (TypeError, ValueError):
        return None
    return None if v != v else v


def pdf_items(path):
    got = []
    PdfReader(str(path)).pages[0].extract_text(
        visitor_text=lambda t, cm, tm, fd, fs: got.append(t.strip()) if t.strip() else None)
    return got


def expected(stem, preset):
    root = ET.parse(SUB / f"{stem}.xml").getroot()
    _, pres = table(root, "Pipes-results")
    _, pin = table(root, "Pipes-input")
    _, nodes = table(root, "Node pressures")

    link = {}
    if preset == "유량본":
        for r in pres:
            f = num(r["Flow"])
            link[key(r["Label"])] = "" if f is None else f"{f * L_MIN:.1f}"
    elif preset == "압력본":
        for r in pres:
            a, b = num(r["Inlet pressure"]), num(r["Outlet pressure"])
            link[key(r["Label"])] = "" if a is None or b is None else f"{(a - b) / KGF_CM2:.1f}"
    else:
        for r in pin:
            d = num(r["Nominal bore"])          # 이미 mm 표기다
            link[key(r["Label"])] = "" if d is None else f"{d:.3f}"

    node = {}
    if preset != "유량본":
        for r in nodes:
            p = num(r["Pressure"])
            node[key(r["Label"])] = "" if p is None else f"{(p - ATM) / KGF_CM2:.1f}"
    return link, node


def check(stem, preset):
    link, node = expected(stem, preset)
    items = pdf_items(OUT / f"{stem}_{preset}.pdf")

    seen_link, seen_node = {}, {}
    for raw in items:
        if "  " in raw:
            label, _, value = raw.partition("  ")
            seen_link[key(label)] = value.strip()
        elif " " in raw:
            label, _, value = raw.partition(" ")
            if key(label) in node:
                seen_node[key(label)] = value.strip()

    bad, absent_link, absent_node = [], [], []
    for label, want in link.items():
        got = seen_link.get(label)
        if got is None:
            absent_link.append(label)
        elif got != want:
            bad.append(f"관로 {label}: PDF {got!r} != XML {want!r}")
    for label, want in node.items():
        got = seen_node.get(label)
        if got is None:
            absent_node.append(label)
        elif got != want:
            bad.append(f"노드 {label}: PDF {got!r} != XML {want!r}")

    print(f"  {preset:<18} 관로 {len(link) - len(absent_link)}/{len(link)} · "
          f"노드 {len(node) - len(absent_node)}/{len(node)} 대조 → 불일치 {len(bad)}"
          f" · 도면에 없는 XML 라벨 관로 {len(absent_link)} 노드 {len(absent_node)}")
    for b in bad[:10]:
        print("      ", b)
    return len(bad), absent_link, absent_node


if __name__ == "__main__":
    total = 0
    for stem in ("2. Pipenet_hand", "2. Pipenet_auto"):
        print("===", stem)
        for preset in ("유량본", "압력본", "압력본_옥내소화전"):
            bad, missing_p, missing_n = check(stem, preset)
            total += bad
        print("   도면에 없는 XML 관로:", missing_p)
        print("   도면에 없는 XML 노드:", missing_n)
    print("총 불일치", total)
    sys.exit(1 if total else 0)
