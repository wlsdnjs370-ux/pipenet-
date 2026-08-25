# -*- coding: utf-8 -*-
"""전체망 .kfp 회귀 기준선 — 이번 작업 전/후로 **비트 동일**해야 한다(지시서 §4).

`.kfp` 는 솔버가 전체망을 받아 설계구역을 스스로 고르는 경로다. 이번 작업은
SDF 쪽에만 손대므로, 이 파일이 한 바이트라도 달라지면 어떤 항목도 완료가 아니다.

    python tests/_kfp_baseline.py make    기준선 생성(작업 전 1회)
    python tests/_kfp_baseline.py check   현재 산출과 대조
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)          # 편집기는 작업 폴더를 cwd 기준으로 잡는다

KEY = "B1F 현장조사 소화설비 평면도"
OUT_DIR = _ROOT / "tests" / "_out"
BASE = _ROOT / "tests" / "_kfp_baseline.json"


def build_kfp() -> tuple[Path, dict]:
    """손질 저장본 → 전체망 .kfp. 기존 경로를 그대로 탄다."""
    from services.cad_import.edit.session import EditSession
    from services.cad_import.convert.engine import convert_to_kfp, ensure_planar
    from services.cad_import.dto import default_dto, dto_to_convert_kwargs

    from services.cad_import.convert.planar import pick_convert_sources

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    # 이 저장본은 급수원이 둘이다. 변환은 하나를 지정해야 하므로 첫 번째(Z1)를
    # 못박는다 — 기준선이므로 «매번 같은 것» 이 값 자체보다 중요하다.
    srcs = payload.get("sources") or ()
    if len(srcs) > 1:
        picked, err = pick_convert_sources(srcs, srcs[0].get("tag"))
        if err:
            raise SystemExit(f"급수원 선택 실패: {err}")
        payload["sources"] = picked
    payload = ensure_planar(payload)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "baseline_full.kfp"
    res = convert_to_kfp(payload, str(out), **dto_to_convert_kwargs(default_dto()))
    if not res["ok"]:
        raise SystemExit(f"변환 실패: {res.get('blockers')}")
    return out, res


def digest(path: Path) -> dict:
    raw = path.read_bytes()
    return {"sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)}


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    out, res = build_kfp()
    cur = digest(out)
    kfp = res["kfp"]
    cur["nodes"] = len(kfp.get("nodes_meta_runtime") or {})
    cur["pipes"] = len(kfp.get("pipe_data") or {})

    if mode == "make":
        BASE.write_text(json.dumps(cur, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        print(f"기준선 생성 · 노드 {cur['nodes']} · 배관 {cur['pipes']} · "
              f"{cur['bytes']:,} bytes\n  sha256 {cur['sha256'][:16]}…")
        return 0

    if not BASE.exists():
        print("!! 기준선이 없다 — 먼저 `make` 로 만들 것")
        return 1
    old = json.loads(BASE.read_text(encoding="utf-8"))
    same = old["sha256"] == cur["sha256"]
    print(f"  기준선 노드 {old['nodes']} 배관 {old['pipes']} {old['bytes']:,}B "
          f"{old['sha256'][:16]}…")
    print(f"  현재   노드 {cur['nodes']} 배관 {cur['pipes']} {cur['bytes']:,}B "
          f"{cur['sha256'][:16]}…")
    print("\n[OK  ] 전체망 .kfp 비트 동일" if same
          else "\n[FAIL] 전체망 .kfp 가 달라졌다 — 어떤 항목도 완료가 아니다")
    return 0 if same else 1


if __name__ == "__main__":
    sys.exit(main())
