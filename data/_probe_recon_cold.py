# -*- coding: utf-8 -*-
"""정찰(덤)이 «처음 보는 도면» 에서 얼마나 무거운가 — 캐시를 건드리지 않고 잰다.

`parse_dxf_bundle_cached` 는 파일 내용 해시로 디스크 캐시를 쓴다. 사용자가
이미 올린 도면은 캐시가 더워져 0.5초로 보이지만, 처음 올리는 도면은 그렇지
않다. 캐시를 **지우지 않고** 비캐시 함수를 직접 불러 그 값을 잰다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

DEFAULT = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> None:
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    print(f"대상: {dxf.name} ({dxf.stat().st_size / 1024 / 1024:,.1f} MB)",
          flush=True)
    import remote30_prototype as A

    t = time.perf_counter()
    key = A._file_content_key(dxf)
    print(f"[a] 내용 해시(캐시 키)     {time.perf_counter() - t:7.2f}s", flush=True)

    cache = A._PARSE_CACHE_DIR / f"v{A._PARSE_CACHE_VERSION}_{key}.pkl.gz"
    print(f"[b] 캐시 있음? {cache.is_file()} "
          f"({cache.stat().st_size / 1024 / 1024:,.1f} MB)"
          if cache.is_file() else "[b] 캐시 없음", flush=True)

    if cache.is_file():
        import gzip
        import pickle
        t = time.perf_counter()
        with gzip.open(cache, "rb") as f:
            pickle.load(f)
        print(f"[c] 캐시 적중 시 비용     {time.perf_counter() - t:7.2f}s",
              flush=True)

    t = time.perf_counter()
    bundle = A.parse_dxf_bundle(dxf)          # 캐시를 쓰지 않는 «처음» 경로
    t_cold = time.perf_counter() - t
    print(f"[d] 캐시 없을 때(처음)    {t_cold:7.2f}s  "
          f"(도형 {len(bundle.entities):,} · 레이어 {len(bundle.layers):,})",
          flush=True)


if __name__ == "__main__":
    main()
