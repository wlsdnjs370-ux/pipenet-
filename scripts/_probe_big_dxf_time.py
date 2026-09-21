# -*- coding: utf-8 -*-
"""큰 도면이 어디서 느린가 — 단계별 실측. 짐작으로 깎지 않기 위한 자.

「느리다」는 한 덩이가 아니다. 열기만 해도 최소 네 토막이다:
    ① DXF 읽기(handoff 캐시 적중이면 건너뜀)  ② 표시 캐시 만들기/읽기
    ③ board 구성(찍은스펙 → 점·간선·헤드)     ④ 화면용 직렬화
그 위에 물흐름·최불리·표 확정이 또 있다. 어느 토막이 절반인지 모르면
최적화는 «안 느린 곳을 빠르게» 만드는 일이 된다.

    python scripts/_probe_big_dxf_time.py --key "..." [--warm|--cold]

--cold 는 표시 캐시를 **옆으로 치우고**(지우지 않는다) 첫 열기를 잰다.
"""
from __future__ import annotations

import argparse
import contextlib
import importlib
import io
import os
import shutil
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
for p in (ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)

BIG = "B1F 현장조사 소화설비 평면도_컨셉2"


class T:
    def __init__(self):
        self.rows = []
        self._t = time.perf_counter()

    def lap(self, label):
        now = time.perf_counter()
        self.rows.append((label, now - self._t))
        self._t = now

    def report(self, title):
        total = sum(d for _l, d in self.rows)
        print(f"\n■ {title} — 합 {total:.1f}s")
        for label, d in self.rows:
            bar = "█" * int(round(d / max(total, 1e-9) * 40))
            print(f"   {d:7.2f}s {d/max(total,1e-9)*100:5.1f}%  {bar} {label}")
        return total


def stage_sizes(key):
    from routes.module_f.common import IMPORT_WORK_ROOT, _boot
    _boot()
    d = str(IMPORT_WORK_ROOT)
    out = {}
    for name, pat in (("표시캐시", f"_edit_disp_cache_{key}.json"),
                      ("찍은스펙", f"0단계_새찍기/{key}_찍은스펙.json")):
        p = os.path.join(d, pat)
        out[name] = os.path.getsize(p) if os.path.isfile(p) else 0
    import glob
    w = glob.glob(os.path.join(d, "0단계_새찍기", "*stage1_world.sqlite3"))
    key_slug = key.replace(" ", "_").replace("(", "").replace(")", "")
    hit = [x for x in w if key_slug[:20].replace(" ", "_") in os.path.basename(x)]
    out["stage1_world"] = os.path.getsize(hit[0]) if hit else 0
    return out


def run(key, cold) -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import IMPORT_WORK_ROOT, _boot
    _boot()
    cache = os.path.join(str(IMPORT_WORK_ROOT),
                         f"_edit_disp_cache_{key}.json")
    moved = None
    if cold and os.path.isfile(cache):
        moved = cache + ".probe_bak"
        shutil.move(cache, moved)
        print(f"  [cold] 표시 캐시를 옆으로 치움 — {os.path.basename(moved)}")

    try:
        sz = stage_sizes(key)
        print(f"■ {key}")
        print("  자산: " + " · ".join(
            f"{k} {v/1048576:.1f}MB" for k, v in sz.items()))

        t = T()
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            from services.cad_import.edit.session import EditSession
            es = EditSession.open(key, out_dir=None, load_saved=True,
                                  use_cache=True)
        t.lap("① EditSession.open (DXF/캐시 → board)")

        b = es.board
        with contextlib.redirect_stdout(buf):
            payload = es.convert_payload()
        t.lap("② convert_payload (변환 입력 만들기)")

        with contextlib.redirect_stdout(buf):
            st = b.water_state()
        t.lap("③ water_state (물흐름)")

        from routes.module_f.views import _edit_state
        sess = {"edit": es, "worst": None, "sheets": [], "id": "probe",
                "key": key}
        with contextlib.redirect_stdout(buf):
            view = _edit_state(sess, full=True)
        t.lap("④ _edit_state (화면용 직렬화)")

        import json as _json
        payload_bytes = len(_json.dumps(view))
        t.lap("⑤ JSON 직렬화 (응답 크기 재기)")

        total = t.report("열기 " + ("(cold — 캐시 없음)" if cold else "(warm)"))
        print(f"   망: 점 {len(b.pts):,} · 간선 {len(b.edges):,} · "
              f"헤드 {len(b.disks):,} · 젖은헤드 {len(st.get('wet_heads') or ()):,}")
        print(f"   화면 응답 {payload_bytes/1048576:.1f} MB")
        for label, d in t.rows:
            if d > total * 0.3:
                print(f"   ★ «{label}» 이 {d/total*100:.0f}% — 여기가 과녁이다")
    finally:
        if moved and os.path.isfile(moved):
            shutil.move(moved, cache)
            print("  [cold] 표시 캐시를 되돌려 놓음")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=BIG)
    ap.add_argument("--cold", action="store_true")
    a = ap.parse_args()
    return run(a.key, a.cold)


if __name__ == "__main__":
    raise SystemExit(main())
