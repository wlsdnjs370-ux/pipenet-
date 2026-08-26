# -*- coding: utf-8 -*-
"""handoff 가 «지문 불일치» 로 캐시를 버리는 이유를 가른다.

캐시 안에는 치수 텍스트가 3,168행(치수로 읽히는 것 533개) 있는데
`load_world` 가 None 을 돌려주면 관경의 도면 텍스트 경로가 통째로 죽는다.
그 판정은 원본 DXF 의 sha256 대조 하나에 달려 있다 — 원본이 지워지거나
다시 올라가면 캐시는 멀쩡한데 못 쓰게 된다.

    python scripts/_probe_handoff_fingerprint.py
"""
from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEY = "B1F 현장조사 소화설비 평면도"


def _sha256(path: str, cap: int = 0) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(1 << 20)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff, stage1 as s1

    # ① 찍은스펙이 가리키는 원본
    spec_path = os.path.join(handoff.pick_out_dir(), f"{KEY}_찍은스펙.json")
    print(f"찍은스펙: {spec_path}")
    print(f"  존재: {os.path.exists(spec_path)}")
    src = None
    if os.path.exists(spec_path):
        with open(spec_path, encoding="utf-8") as f:
            src = json.load(f).get("source_dxf")
    print(f"  source_dxf: {src}")
    print(f"  그 파일 존재: {os.path.exists(src) if src else False}")

    # ② 캐시가 기억하는 원본
    hp = handoff.handoff_path(KEY)
    con = sqlite3.connect(hp)
    try:
        meta = dict(con.execute("SELECT key,value FROM meta").fetchall())
    finally:
        con.close()
    print(f"\n캐시가 기억하는 원본")
    print(f"  source_path : {meta.get('source_path')}")
    print(f"  source_size : {int(meta.get('source_size', 0)):,}")
    print(f"  source_sha  : {meta.get('source_sha256')}")
    print(f"  n_texts     : {meta.get('n_texts')}")

    # ③ 실제 파일과 대조
    if src and os.path.exists(src):
        size = os.path.getsize(src)
        print(f"\n실제 파일")
        print(f"  size : {size:,}  ({'같음' if str(size) == meta.get('source_size') else '★다름'})")
        print("  sha256 계산 중…")
        got = _sha256(src)
        same = got == meta.get("source_sha256")
        print(f"  sha  : {got}  ({'같음' if same else '★다름'})")
    else:
        print("\n★ 찍은스펙이 가리키는 원본 DXF 가 없다 — 지문 대조 자체가 불가능")

    # ③-b _meta_matches 의 6개 조건을 하나씩 —— 어느 것이 튕기나
    if src and os.path.exists(src):
        from services.cad_import.pipeline.handoff import (
            FORMAT, _compatible_prep_digest, _sha256_file)
        norm = os.path.normcase(os.path.abspath(src))
        st = os.stat(norm)
        conds = [
            ("format", meta.get("format"), FORMAT),
            ("prep_sha256", meta.get("prep_sha256"), _compatible_prep_digest()),
            ("source_path", meta.get("source_path"), norm),
            ("source_size", meta.get("source_size"), str(st.st_size)),
            ("source_mtime_ns", meta.get("source_mtime_ns"), str(st.st_mtime_ns)),
            ("source_sha256", meta.get("source_sha256"), _sha256_file(norm)),
        ]
        print("\n_meta_matches 조건별")
        for name, got, want in conds:
            ok = got == want
            print(f"  [{'OK  ' if ok else 'FAIL'}] {name}")
            if not ok:
                print(f"          캐시: {got}")
                print(f"          현재: {want}")

    # ④ load_world 가 실제로 무엇을 돌려주나
    w = handoff.load_world(KEY, src, s1.World) if src else None
    print(f"\nload_world 결과: {'World' if w is not None else 'None ★'}")
    if w is not None:
        print(f"  texts {len(w.texts):,} · segs {len(w.segs):,}")
        from services.cad_import.design.bore import extract_dia_text_points
        print(f"  치수로 읽힌 것 {len(extract_dia_text_points(w.texts)):,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
