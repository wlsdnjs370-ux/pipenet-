# -*- coding: utf-8 -*-
"""[§29] 알람밸브 배선 전/후 — 산출물이 얼마나 달라지는가.

D-F11-1 은 「사람의 명시적 수정 외에는 산출을 바꾸지 않는다」이다. 이 배선은
산출을 바꾸므로, 얼마나 바뀌는지를 먼저 재고 기록한다.
"""
from __future__ import annotations

import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

from routes.module_f.common import _boot                           # noqa: E402
_boot()
from services.cad_import.design.anchor import valve_kfp_nodes      # noqa: E402
from services.cad_import.design.restrict import select_and_expand  # noqa: E402
from services.cad_import.design.tables import build_design_tables  # noqa: E402
from services.cad_import.design.worst import worst_k_heads         # noqa: E402
from services.cad_import.edit.session import EditSession, MODE_VALVE  # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
print(f"손질 상태 — 알람밸브 {list(b.valves)} · 접속점 {list(b.sources)}")

# 옛 저장본은 급수원만 있고 밸브가 없다(리팩터링 7 이전에 찍은 것).
# 사람이 지금 화면에서 하듯 «접속점 자리에» 알람밸브를 찍어 준다.
if not b.valves and b.sources:
    n = b.sources[0]
    es.set_mode(MODE_VALVE)
    rep = es.click(float(b.pts[n][0]), float(b.pts[n][1]), 500.0)
    print(f"  → 알람밸브를 접속점 자리에 찍었다: {rep} · "
          f"밸브 {list(b.valves)} · 접속점 {list(b.sources)}")

w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)
got = select_and_expand(es.convert_payload(), b, k=30)
if not got.get("ok"):
    raise SystemExit(f"전개 실패: {got.get('error')}")

kfp = got["kfp"]["nodes_meta_runtime"]
av, missed = valve_kfp_nodes(kfp, b.pts, list(b.valves), got.get("origin_mm"))
print(f"\n되짚기 — 밸브 {len(b.valves)}곳 → kfp 노드 {av} · 못 이은 것 {missed}")
if av:
    m = kfp[av[0]]
    print(f"  그 노드 type_id={m.get('type_id')!r} "
          f"coords={[round(float(c), 3) for c in (m.get('coords') or [])]}")

common = dict(board_pts=b.pts, tree_loads=got.get("tree_loads"),
              origin_mm=got.get("origin_mm"))
before = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                             valve_nodes=None, **common)
after = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                            valve_nodes=av, **common)

print("\n■ 산출물 차이")
print(f"    기기표 행  {len(before.equipment)} → {len(after.equipment)}")
dia_of = {str(r["label"]): r.get("dia") for r in after.pipes}
for e in after.equipment:
    print(f"      {e['desc']} · 배관 {e['pipe']} ({e['in']}→{e['out']})"
          f" · 호칭경 {dia_of.get(str(e['pipe']))}A"
          f" · 등가길이 {e['eq_len']} m · 근거 {e.get('eq_len_src')}")
from services.cad_import.design.fitting import load_equivalent_lengths
_lib = load_equivalent_lengths()
print(f"      라이브러리 VALVE_ALARM 이 가진 호칭경: "
      f"{sorted(_lib.get('VALVE_ALARM', {}))}")
print("      배관 앞 8개 — 라벨 · 관경 · 근거 · 담당헤드")
loads = got.get("tree_loads") or {}
for r in after.pipes[:8]:
    print(f"        {r['label']:>6s} {str(r.get('dia')):>4s}A "
          f"src={r.get('dia_src')!r:>12s} load={loads.get(str(r['label']))}"
          f"  {r['in']}→{r['out']} len={r.get('length')}")
for name in ("nodes", "pipes", "nozzles", "fittings"):
    a, c = len(getattr(before, name)), len(getattr(after, name))
    print(f"    {name:9s} {a} → {c}" + ("  ← 안 바뀜" if a == c else "  ★바뀜"))
print(f"    등가길이 미해결 {dict(before.meta)['등가길이 미해결']} → "
      f"{dict(after.meta)['등가길이 미해결']}")

# 배관 총연장 대비 얼마나 되는가 — §29 가 잰 그 비율.
tot = sum(float(r.get("length") or 0.0) for r in after.pipes)
eq = sum(float(r.get("eq_len") or 0.0) for r in after.equipment)
print(f"\n    배관 총연장 {tot:.1f} m · 기기 등가길이 {eq:.1f} m "
      f"({eq / max(tot, 1e-9) * 100:.1f}%)")
print("    ※ §29 가 적은 «FX 32개 × 15.6m» 는 아직 없다 — 이번 배선은"
      " 알람밸브뿐이다.")

# lift 영점도 함께 산다(§29 곁따라).
from routes.module_f.api_design import _valve_label                # noqa: E402
print(f"\n    lift 영점 «알람밸브» 라벨: {_valve_label(before)!r} → "
      f"{_valve_label(after)!r}")
