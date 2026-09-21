# -*- coding: utf-8 -*-
import io
p = "scripts/_verify_module_f_browser.py"
s = io.open(p, encoding="utf-8").read()
old = '''    w = page.evaluate("() => window.__mf.edit.worst")
    print(f"   헤드 {w['k']} · 경로 {len(w['path'])} 간선 ·"
          f" 최원 {w['far_m']} m · 끝 {w['near_m']} m")
    if w["k"] != 30 or not w["path"]:
        bad(f"최불리 선정 결과가 비었다: {w}")'''
new = '''    w = page.evaluate("() => window.__mf.edit.worst")
    print(f"   설계면적 {w['k']}개 · corridor {len(w['corridor'])} 간선 ·"
          f" 앵커 {w['far_m']} m · 폭 {w['span_m']} m ·"
          f" 연장 {w['total_m']} m · 주배관 {w['max_load']}개 담당")
    if w["k"] != 30 or not w["corridor"]:
        bad(f"최불리망 결과가 비었다: k={w['k']} corridor={len(w.get('corridor', []))}")
    # 앵커(가장 불리한 지점)가 실려 화면에 그려져야 한다.
    if not w.get("anchor"):
        bad("앵커 헤드가 실리지 않았다")
    # corridor 간선마다 담당 헤드 수(load)가 붙고, 주배관은 여러 개를 먹인다.
    loads = [c[4] for c in w["corridor"]]
    if not loads or max(loads) != w["max_load"] or max(loads) < 2:
        bad(f"담당 헤드 수(load)가 corridor 에 안 실렸다: max={w['max_load']}")
    print(f"   load 분포: 최대 {max(loads)} · load=1 가지 {sum(1 for x in loads if x == 1)}개")'''
assert old in s
s = s.replace(old, new, 1)
old2 = '''    if after - before < 400:
        bad(f"최불리 경로가 캔버스에 그려지지 않았다 ({before} → {after})")'''
new2 = '''    if after - before < 400:
        bad(f"최불리망이 캔버스에 그려지지 않았다 ({before} → {after})")'''
assert old2 in s
s = s.replace(old2, new2, 1)
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("브라우저 검증 [6-A] 를 corridor/앵커/load 로 갱신")
