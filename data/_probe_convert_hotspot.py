# -*- coding: utf-8 -*-
"""전체망 변환의 «14분 침묵» 이 어디에 있나 — 프로파일로 먼저 본다.

노드 5,070 · 배관 5,069 짜리 .kfp 를 쓰는 데 872초는 규모에 안 맞는다. 말없이
오래 걸리는 것을 그냥 말하게만 하면, 사실은 고칠 수 있었던 것을 «원래 오래
걸리는 일» 로 굳혀 버린다. 그래서 narration 을 붙이기 «전에» 어디서 타는지
잰다.

실행: python data/_probe_convert_hotspot.py
"""
from __future__ import annotations

import cProfile
import io
import os
import pstats
import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

KEY = "B1F 현장조사 소화설비 평면도"


def main() -> None:
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.convert.engine import (
        convert_to_kfp, ensure_planar)
    from services.cad_import.dto import default_dto, dto_to_convert_kwargs
    from services.cad_import.edit.session import EditSession

    t = time.perf_counter()
    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    print(f"[1] 손질 열기      {time.perf_counter() - t:7.2f}s", flush=True)

    t = time.perf_counter()
    payload = es.convert_payload()
    print(f"[2] convert_payload {time.perf_counter() - t:7.2f}s", flush=True)

    out = ROOT / "data" / "uploads" / "module_f" / "_hotspot_probe.kfp"
    out.parent.mkdir(parents=True, exist_ok=True)
    kw = dto_to_convert_kwargs(default_dto())

    # ★첫 시도는 `convert_to_kfp` 만 감쌌는데 그 안이 아니었다. 실측 로그를
    #   다시 읽으니 「전체망 — 수직 전개 후 .kfp 를 씁니다…」 줄이 **869초에**
    #   비로소 찍혔다 — 침묵은 그 앞, `ensure_planar`(평면 그래프 만들기)에
    #   있었다. 둘 다 한 프로파일에 넣고 시간도 따로 잰다.
    print("[3] ensure_planar + convert_to_kfp — 프로파일 시작", flush=True)
    pr = cProfile.Profile()
    pr.enable()
    t = time.perf_counter()
    payload = ensure_planar(payload)
    t_planar = time.perf_counter() - t
    t = time.perf_counter()
    res = convert_to_kfp(payload, str(out), **kw)
    t_conv = time.perf_counter() - t
    pr.disable()
    print(f"[3] ensure_planar   {t_planar:7.2f}s", flush=True)
    print(f"[4] convert_to_kfp  {t_conv:7.2f}s · ok={res['ok']}", flush=True)

    buf = io.StringIO()
    st = pstats.Stats(pr, stream=buf)
    st.sort_stats("tottime").print_stats(22)
    print(buf.getvalue()[:6000], flush=True)


if __name__ == "__main__":
    main()
