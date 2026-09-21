# -*- coding: utf-8 -*-
"""모듈 F «열기 잡» 구간 실측 — 올린 뒤 도면이 뜨기까지 무엇이 시간을 먹나.

`_open_job` 은 세 토막이다:

    (1) PickSession.open      도면을 읽어 찍기판을 세운다   ← 여기까지가 «도면»
    (2) _world_payload        캔버스로 내려보낼 모양 만들기 ← 여기까지가 «도면»
    (3) _recon_into           A 로 «한 번 더» 읽어 헤드 후보를 뽑는다 (덤)

화면은 잡이 «끝난 뒤에» 도면을 그린다(module_f.js 의 watch → onDone). 그래서
(3) 이 길면 도면이 그만큼 늦게 뜬다 — (3) 은 덤인데 값은 앞에서 치른다.

실행: python data/_probe_module_f_openjob.py [dxf경로]
"""
from __future__ import annotations

import io
import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

DEFAULT = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> None:
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not dxf.is_file():
        raise SystemExit(f"파일 없음: {dxf}")
    mb = dxf.stat().st_size / 1024 / 1024
    print(f"대상: {dxf.name} ({mb:,.1f} MB)", flush=True)

    from routes.module_f.common import _boot
    t = time.perf_counter()
    _boot()
    print(f"[0] _boot                {time.perf_counter() - t:7.2f}s", flush=True)

    from routes.module_f.world import _world_payload
    from services.cad_import.pick.session import PickSession

    t = time.perf_counter()
    ps = PickSession.open(str(dxf))
    t_open = time.perf_counter() - t
    print(f"[1] PickSession.open     {t_open:7.2f}s", flush=True)

    t = time.perf_counter()
    ps.select_pipe()
    payload = _world_payload(ps.world)
    t_pay = time.perf_counter() - t
    c = payload["counts"]
    print(f"[2] _world_payload       {t_pay:7.2f}s  "
          f"(선분 {c['segs']:,} · 원 {c['circles']:,} · 호 {c['arcs']:,})",
          flush=True)

    # 화면으로 실제 내려가는 JSON 크기 — /api/module-f/world 의 몸통.
    import json
    t = time.perf_counter()
    blob = json.dumps({"ok": True, "world": payload}, ensure_ascii=False)
    t_json = time.perf_counter() - t
    print(f"[2b] world JSON 직렬화   {t_json:7.2f}s  "
          f"({len(blob.encode('utf-8')) / 1024 / 1024:,.1f} MB)", flush=True)
    del blob

    t = time.perf_counter()
    from routes.module_f.recon import run_recon
    try:
        rec = run_recon(dxf, world=payload)
        t_rec = time.perf_counter() - t
        print(f"[3] _recon_into (덤)     {t_rec:7.2f}s  "
              f"(후보 {len(rec['heads']):,})", flush=True)
    except BaseException as exc:  # noqa: BLE001
        t_rec = time.perf_counter() - t
        print(f"[3] _recon_into 실패     {t_rec:7.2f}s  "
              f"{type(exc).__name__}: {exc}", flush=True)

    total = t_open + t_pay + t_rec
    print()
    print(f"도면이 뜰 수 있는 시점   {t_open + t_pay:7.2f}s")
    print(f"실제로 뜨는 시점(현행)   {total:7.2f}s   "
          f"← 덤이 {t_rec / max(total, 1e-9) * 100:.0f}% 를 차지")


if __name__ == "__main__":
    main()
