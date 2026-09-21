# -*- coding: utf-8 -*-
"""SDF 가 적어 둔 글자 크기 값이 실제 도면의 글씨 크기를 설명하는지 본다.

렌더러는 typesize 의 단위를 모른다고 보고 상대비만 지켜 왔다. 실측한
모델 단위 글자 높이를 SDF 값으로 나눠 그 몫이 도면마다 같으면, 단위를
모르는 게 아니라 모델 단위인 것이다.
"""
from __future__ import annotations

import collections
import pathlib
import re
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"
NUM = re.compile(r"^[-+]?\d[\d.,]*$")


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    keys = collections.Counter()
    rows = []
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                  # noqa: BLE001
            continue
        doc = fitz.open(pdf)
        page = doc[0]
        W, H = page.rect.width, page.rect.height
        dots = [d["rect"] for d in page.get_drawings()
                if "".join(it[0] for it in d["items"]) == "cccc"
                and d["fill"] is not None and max(d["fill"]) < 0.01]
        spans = [s for b in page.get_text("dict")["blocks"] for ln in b.get("lines", ())
                 for s in ln["spans"]]
        doc.close()
        real = [n for n in model.nodes if not n.virtual]
        if W > H or len(dots) < 4 or len(real) < 4 or abs(len(real) - len(dots)) > 1:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        dspan = max(max(r.x0 for r in dots) - min(r.x0 for r in dots),
                    max(r.y0 for r in dots) - min(r.y0 for r in dots))
        sizes = [s["size"] for s in spans
                 if 0.156 < 1 - s["bbox"][3] / H and 1 - s["bbox"][3] / H < 0.930
                 and NUM.match(s["text"].strip())]
        if mspan <= 0 or dspan <= 0 or len(sizes) < 10:
            continue
        size_model = statistics.median(sizes) / (dspan / mspan)
        lab = model.display_options.get("Label-display", {})
        for k in lab:
            keys[k] += 1
        rows.append((sdf.stem, size_model, lab))
        if len(rows) >= limit:
            break

    print(f"짝 {len(rows)} 세트")
    print(f"  Label-display 속성: {keys.most_common()}")
    # 숫자로 읽히는 속성마다, 실측 높이를 그 값으로 나눈 몫이 얼마나 고른지 본다.
    names = {k for _, _, lab in rows for k, v in lab.items()
             if re.match(r"^[\d.]+$", v or "")}
    for name in sorted(names):
        pairs = [(sm, float(lab[name])) for _, sm, lab in rows
                 if lab.get(name) and float(lab[name]) > 0]
        if len(pairs) < len(rows) * 0.8:
            continue
        vals = sorted(sm / f for sm, f in pairs)
        raw = sorted({f for _, f in pairs})
        print(f"  {name:16s} 값 {raw[:6]}  몫 중앙 {statistics.median(vals):.4f}"
              f"  변동계수 {statistics.pstdev(vals) / statistics.mean(vals):.3f}")
        if len(raw) > 1:
            # 값이 갈리는 속성은 그 값별로 실측 높이가 정말 갈리는지 따로 본다.
            by = collections.defaultdict(list)
            for sm, f in pairs:
                by[f].append(sm)
            for f in sorted(by):
                g = sorted(by[f])
                print(f"      {name}={f:g} 인 {len(g):2d}장: 실측 높이 중앙 "
                      f"{statistics.median(g):.3f} (최소 {g[0]:.3f} 최대 {g[-1]:.3f})")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
