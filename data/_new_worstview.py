NEW = '''def _worst_view(sess: dict) -> dict | None:
    """화면용 — 최불리 배관망(corridor)·앵커·담당 헤드 수.

    corridor 간선은 좌표 4개 + load 를 함께 싣는다(화면이 굵기/색을 정한다).
    앵커는 «가장 불리한 지점» 이라 따로 강조한다.
    """
    w = sess.get("worst")
    if not w:
        return None
    b = sess["edit"].board
    pts = b.pts
    disks = b.disks
    an = w.get("anchor")
    return {
        "k": len(w["heads"]),
        "reachable": w["reachable"],
        "far_m": w["far_m"],
        "near_m": w["near_m"],
        "span_m": w.get("span_m", 0.0),
        "total_m": w.get("total_m", 0.0),
        "max_load": w.get("max_load", 0),
        "sheet": w.get("sheet"),
        "heads": [[_r1(disks[hi][0]), _r1(disks[hi][1]), _r1(disks[hi][2])]
                  for hi in w["heads"] if hi < len(disks)],
        "anchor": ([_r1(disks[an][0]), _r1(disks[an][1]), _r1(disks[an][2])]
                   if isinstance(an, int) and an < len(disks) else None),
        "corridor": [[_r1(pts[a][0]), _r1(pts[a][1]),
                      _r1(pts[c][0]), _r1(pts[c][1]), int(load)]
                     for (a, c), load in w.get("loads", {}).items()],
    }'''
import io
p = "routes/module_f/remote30.py"
s = io.open(p, encoding="utf-8").read()
lo = s.index("def _worst_view(")
hi = s.index("# ─", lo)   # 다음 구분선(자동 이음 헤더) 앞까지
s = s[:lo] + NEW + "\n\n\n" + s[hi:]
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("_worst_view 교체 완료")
