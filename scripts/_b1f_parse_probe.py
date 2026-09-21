# -*- coding: utf-8 -*-
"""B1F(120MB) 파싱 병목 실측 — 추출 경로(parse_dxf_bundle)의 비용 구조 진단.

측정:
  1) ezdxf.readfile 시간
  2) msp top-level entity 수 + 카테고리(전경 PIPE/HEAD/TEXT/ALARM vs 배경 ARCH/EXCLUDE/OTHER) 분리
  3) top-level INSERT 폭발 leaf 추정(memo, 렌더 없이) — 전경/배경 각각
  4) 배경 top-level 안에 nested 된 PIPE/HEAD leaf 수(LH 병리: top-level 분리가 안 통하는 구조인지)
  5) parse_dxf_bundle 전체 시간 + 산출 entity 카테고리 분포 + filter_pipenet_only 잔량
"""
from __future__ import annotations

import sys
import time
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

import ezdxf  # noqa: E402
from remote30_prototype import (  # noqa: E402
    _categorize_layer, parse_dxf_bundle, filter_pipenet_only,
)

DXF = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
MAX_DEPTH = 10
FG_CATS = {"PIPE", "HEAD", "TEXT", "ALARM"}


def _leaf_breakdown(block_name, depth, doc, memo):
    """블록 폭발 시 leaf 를 카테고리별 Counter 로 memo. (블록명,depth) 키."""
    if depth >= MAX_DEPTH or doc is None:
        return Counter()
    key = (block_name, depth)
    if key in memo:
        return memo[key]
    memo[key] = Counter()  # 순환참조 가드
    blk = doc.blocks.get(block_name)
    acc: Counter = Counter()
    if blk is not None:
        for child in blk:
            if child.dxftype() == "INSERT":
                acc += _leaf_breakdown(child.dxf.name, depth + 1, doc, memo)
            else:
                ly = str(getattr(child.dxf, "layer", ""))
                acc[_categorize_layer(ly)] += 1
    memo[key] = acc
    return acc


def main():
    assert DXF.is_file(), f"DXF 없음: {DXF}"
    print(f"file: {DXF.name}  ({DXF.stat().st_size/1e6:.1f} MB)")

    t = time.time()
    doc = ezdxf.readfile(str(DXF))
    msp = doc.modelspace()
    print(f"[1] ezdxf.readfile: {time.time()-t:.1f}s")

    top = list(msp)
    print(f"[2] top-level entities: {len(top):,}")

    memo: dict = {}
    fg_top = bg_top = 0
    fg_leaf: Counter = Counter()
    bg_leaf: Counter = Counter()
    t = time.time()
    for e in top:
        et = e.dxftype()
        ly = str(getattr(e.dxf, "layer", ""))
        cat = _categorize_layer(ly)
        is_bg_top = cat in ("ARCH", "EXCLUDE", "OTHER")
        if is_bg_top:
            bg_top += 1
        else:
            fg_top += 1
        if et == "INSERT":
            br = _leaf_breakdown(e.dxf.name, 0, e.doc, memo)
        else:
            br = Counter({cat: 1})
        (bg_leaf if is_bg_top else fg_leaf).update(br)
    print(f"[3] top-level split (배경=ARCH/EXCLUDE/OTHER): 전경 {fg_top:,} / 배경 {bg_top:,}"
          f"  (leaf estimate {time.time()-t:.1f}s)")
    print(f"    전경 top 안 leaf: {dict(fg_leaf)}  (합 {sum(fg_leaf.values()):,})")
    print(f"    배경 top 안 leaf: {dict(bg_leaf)}  (합 {sum(bg_leaf.values()):,})")
    buried = bg_leaf.get("PIPE", 0) + bg_leaf.get("HEAD", 0) + bg_leaf.get("ALARM", 0)
    print(f"[4] 배경 top-level 안에 nested PIPE/HEAD/ALARM leaf: {buried:,}  "
          f"(>0 이면 top-level 분리로는 배관망 손실 위험=LH 병리)")

    t = time.time()
    bundle = parse_dxf_bundle(DXF)
    print(f"[5] parse_dxf_bundle 전체: {time.time()-t:.1f}s  → {len(bundle.entities):,} entity")
    cat_dist: Counter = Counter(_categorize_layer(en["l"]) for en in bundle.entities)
    print(f"    산출 entity 카테고리 분포: {dict(cat_dist)}")
    t = time.time()
    pipe_ents = filter_pipenet_only(bundle)
    print(f"[6] filter_pipenet_only: {time.time()-t:.1f}s  → {len(pipe_ents):,} entity "
          f"({100*len(pipe_ents)/max(len(bundle.entities),1):.1f}% 유지)")
    print("DONE")


if __name__ == "__main__":
    main()
