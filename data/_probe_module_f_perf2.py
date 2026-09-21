# -*- coding: utf-8 -*-
"""망 지문 게이트 전/후 비교 — 클릭 한 번의 무게."""
from __future__ import annotations
import json, os, sys, time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT); sys.path.insert(0, ROOT)
import routes.module_f as mf  # noqa: E402
mf._boot(); sys.path.append(str(mf.EDITOR_ROOT))
from services.cad_import.edit.session import EditSession  # noqa: E402

es = EditSession.open("B1F 현장조사 소화설비 평면도", out_dir=None,
                      load_saved=True, use_cache=True)
b = es.board
sess = {"id": "x", "key": "k", "edit": es, "water_path": None, "worst": None,
        "autojoin": None, "autojoin_report": None, "sheets": None}

def one(label, **kw):
    t0 = time.perf_counter()
    st = mf._edit_state(sess, **kw)
    ms = (time.perf_counter() - t0) * 1000
    kb = len(json.dumps(st)) / 1024
    print(f"  {label:<40} {ms:7.1f} ms · {kb:7.0f} KB · 덩이 {st['counts']['bodies']}"
          f" · 도형 {len(st['body_groups'])}묶음")
    return st

print("[화면 첫 조회 — 반드시 전량]")
one("edit/state (full=True)", full=True)
print("\n[망이 안 바뀐 클릭 — 이음 첫 클릭·헤드선택·급수토글]")
one("edit/click (net=True, 변화 없음)")
one("edit/click (net=True, 변화 없음)")
print("\n[모드 전환]")
one("edit/mode (지문이 같으면 자동으로 가벼움)")

print("\n[망이 바뀐 클릭 — 실제 이음]")
segs = b.segments()
made = 0
for i in range(0, len(segs), 977):
    for j in range(i + 1, min(i + 40, len(segs))):
        n, _bl, _cv, _k = b.join(segs[i], segs[j])
        if n:
            made += 1
            break
    if made:
        break
print(f"  (테스트용 이음 {made}곳)")
one("edit/click (net=True, 변화 있음)")
one("edit/click (net=True, 다시 — 변화 없음)")

print("\n[급수원 토글 — 도형은 그대로, 닿는 헤드는 바뀌어야 한다]")
st = mf._edit_state(sess, full=True)
print(f"  전: source_heads={st['body_stat']['source_heads']}"
      f" has_source={st['body_stat']['has_source']}")
for n in list(b.sources):
    b.remove_source(n)
st = mf._edit_state(sess)
print(f"  급수원 해제 후: source_heads={st['body_stat']['source_heads']}"
      f" has_source={st['body_stat']['has_source']}  keep={st['keep']}")
