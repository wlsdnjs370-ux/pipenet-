# -*- coding: utf-8 -*-
import io
p = "scripts/_verify_module_f.py"
s = io.open(p, encoding="utf-8").read()
old = '''            check("경로 간선 축소", 0 < s["path_edges"] < st["counts"]["edges"],
                  f"{s['path_edges']} / 전체 {st['counts']['edges']}")
            w = j["state"]["worst"]
            check("화면 좌표 동봉", len(w["heads"]) == 30 and len(w["path"]) > 0,
                  f"헤드 {len(w['heads'])} · 경로 {len(w['path'])}")'''
new = '''            check("경로 간선 축소", 0 < s["path_edges"] < st["counts"]["edges"],
                  f"{s['path_edges']} / 전체 {st['counts']['edges']}")
            w = j["state"]["worst"]
            check("최불리망 좌표 동봉",
                  len(w["heads"]) == 30 and len(w["corridor"]) > 0,
                  f"헤드 {len(w['heads'])} · corridor {len(w['corridor'])}")
            # ★설계면적 — 앵커 방식은 «먼 순서» 와 달리 헤드가 뭉쳐야 한다.
            #   퍼짐(대각)을 재서 배관 연장보다 한참 작은지 본다.
            import math as _m
            hp = [(h[0], h[1]) for h in w["heads"]]
            xs = [p_[0] for p_ in hp]; ys = [p_[1] for p_ in hp]
            diag_m = _m.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0
            check("설계면적으로 뭉침(퍼짐 < 총연장)",
                  diag_m < s["total_m"],
                  f"퍼짐 {diag_m:.1f} m · 배관연장 {s['total_m']} m · 폭 {s['span_m']} m")
            check("앵커 = 최원 유하거리", w.get("anchor") is not None
                  and abs(s["far_m"] - w["far_m"]) < 0.01,
                  f"앵커 {w.get('anchor')} · 최원 {s['far_m']} m")
            # 담당 헤드 수 — 주배관은 여러 헤드를 먹이고, 말단 가지는 load=1.
            loads = [c_[4] for c_ in w["corridor"]]
            check("담당 헤드 수(load) 실림",
                  s["max_load"] >= 1 and s["max_load"] == max(loads),
                  f"최대 {s['max_load']} · load=1 가지 {sum(1 for x in loads if x == 1)}개")'''
assert old in s
s = s.replace(old, new, 1)
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("검증에 설계면적·앵커·load 검사 추가")
