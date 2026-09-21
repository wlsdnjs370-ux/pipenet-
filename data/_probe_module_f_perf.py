# -*- coding: utf-8 -*-
"""모듈 F 손질 단계 병목 실측 — 무엇이 느린지 재고 나서 고친다."""
from __future__ import annotations

import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
sys.path.insert(0, ROOT)

import routes.module_f as mf  # noqa: E402

mf._boot()
sys.path.append(str(mf.EDITOR_ROOT))
from services.cad_import.edit.session import EditSession  # noqa: E402

KEY = "B1F 현장조사 소화설비 평면도"


def t(label, fn, n=3):
    fn()  # warm
    best = min((lambda: (lambda s: (fn(), time.perf_counter() - s)[1])(time.perf_counter()))()
               for _ in range(n))
    print(f"  {label:<34} {best*1000:8.1f} ms")
    return best


t0 = time.perf_counter()
es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
print(f"열기 {time.perf_counter()-t0:.1f}s · 노드 {len(es.board.pts)} · "
      f"간선 {len(es.board.edges)} · 헤드 {len(es.board.disks)}")
b = es.board
sess = {"id": "x", "edit": es, "water_path": None, "worst": None,
        "autojoin": None, "autojoin_report": None, "sheets": None}

print("\n[단품]")
t("board.bodies()", lambda: b.bodies())
t("_body_index", lambda: mf._body_index(b))
t("_body_stat", lambda: mf._body_stat(b))
t("display_geom(net=True)", lambda: es.display_geom(net=True))
t("display_geom(net=False)", lambda: es.display_geom(net=False))
t("_sheet_frames", lambda: mf._sheet_frames(b), n=2)
t("_autojoin_scan", lambda: mf._autojoin_scan(b), n=2)

print("\n[상태 한 번 = 클릭 한 번]")
t("_edit_state(net=True)", lambda: mf._edit_state(sess), n=3)
t("_edit_state(net=False)", lambda: mf._edit_state(sess), n=3)

st = mf._edit_state(sess)
raw = json.dumps(st)
print(f"\n상태 JSON {len(raw)/1024:.0f} KB")
parts = sorted(((len(json.dumps(v)) / 1024, k) for k, v in st.items()),
               reverse=True)[:8]
for kb, k in parts:
    print(f"  {k:<20} {kb:8.0f} KB")

sess["autojoin"] = mf._autojoin_scan(b)
st2 = mf._edit_state(sess)
print(f"\n후보 표시 중 상태 JSON {len(json.dumps(st2))/1024:.0f} KB "
      f"(autojoin {len(json.dumps(st2['autojoin']))/1024:.0f} KB)")
