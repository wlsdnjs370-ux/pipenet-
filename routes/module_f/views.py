# -*- coding: utf-8 -*-
"""화면이 받는 상태 한 장.

★망 도형은 «바뀌었을 때만» 싣는다. 무엇이 안 바뀌었는지는 `keep` 으로 알린다 —
빈 배열만으로는 «비었다» 와 구별이 안 된다.
"""
from __future__ import annotations

from routes.module_f.common import _r1
from routes.module_f.graph import _body_stat
from routes.module_f.remote30 import _worst_view
from routes.module_f.world import _pts_bounds

def _pick_state(sess: dict) -> dict:
    ps = sess["pick"]
    hl = ps.highlight_geom()
    spec = ps.spec()
    return {
        "mode": ps.mode,
        "armed": ps.armed,
        "mat_done": bool(ps.mat_done),
        "head_label": ps.head_label,
        "materials": [{"layer": ly, "color": c} for ly, c in ps.board.mat],
        "n_heads": len(ps.board.heads),
        "n_clicks": len(ps.board.clicks),
        "clicks": ps.board.clicks[-12:],
        "highlight": {
            "pipe_bundles": [f"{ly}{c}" for ly, c in hl["pipe_bundles"]],
            "pipe_segs": [[_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
                          for a, b in hl["pipe_segs"]],
            "head_circles": [[_r1(x), _r1(y), _r1(r)]
                             for x, y, r in hl["head_circles"]],
            "tri_segs": [[_r1(a[0]), _r1(a[1]), _r1(b[0]), _r1(b[1])]
                         for a, b in hl["tri_segs"]],
            "last_click": hl["last_click"],
        },
        "n_spec_heads": len(spec.get("heads") or []),
        "n_ho": len(spec.get("ho") or []),
    }


def _net_rev(board) -> tuple:
    """망이 바뀌었는지 O(1) 로 가리는 지문.

    이음은 간선(T분기는 노드까지)을 늘리고 삭제는 줄인다. 길이 넷이면 구조
    변화는 전부 잡힌다 — 좌표만 바뀌는 연산은 이 판에 없다.
    """
    return (len(board.pts), len(board.edges),
            len(board.joins), len(board.deletes))


def _edit_state(sess: dict, full: bool = False) -> dict:
    """손질 화면 한 장.

    ★망 도형은 **바뀌었을 때만** 싣는다. 실측 B1F 는 덩이 선분이 19,797개라
    한 장이 935KB·126ms 인데, 이음 첫 클릭·헤드 선택·급수 토글처럼 **망이 안
    바뀌는 클릭이 절반**이다. 안 바뀌었으면 `body_groups` 를 빈 배열로 두고
    화면이 들고 있던 것을 그대로 쓰게 한다(모드 전환이 이미 쓰던 규약).
    `full=True` 는 화면이 제 사본을 잃었을 때(첫 조회·새로고침) 강제 전송.

    ★«망 재계산 생략(net=False)» 스위치는 두지 않는다. 예전에는 모드 전환만
    가벼우라고 그 인자를 받았는데, 그러면 «망은 바뀌었는데 net=False 로 부른»
    조합에서 낡은 사본을 지키라고 시키게 된다. 지문 하나면 두 목적을 다 이룬다.
    """
    from services.cad_import.colors import (
        EDIT_SOURCE, EDIT_VALVE, EDIT_WET_PIPE, KIND_COLORS)
    es = sess["edit"]
    b = es.board
    rev = _net_rev(b)
    keep: list[str] = []

    net_fresh = full or sess.get("net_rev") != rev
    g = es.display_geom(net=net_fresh)
    if net_fresh:
        sess["net_rev"] = rev
        sess["n_bodies"] = len(g["body_groups"])
    else:
        keep.append("body_groups")

    # ★급수원·밸브는 도형을 안 바꾸므로 망 지문에 안 걸린다. 그런데 «급수원이
    # 닿는 헤드» 는 바뀐다 — 도형 지문으로 같이 묶으면 급수원을 찍어도 0 으로
    # 남는다. 그래서 통계는 제 지문을 따로 갖는다.
    stat_rev = rev + (len(b.sources), len(b.valves))
    if full or sess.get("stat_rev") != stat_rev or not sess.get("body_stat"):
        sess["stat_rev"] = stat_rev
        sess["body_stat"] = _body_stat(b)

    # 헤드 색은 종류·젖음으로만 바뀐다. 지문은 «정확히» 잡는다 — 개수만 세면
    # 이미 지정된 헤드를 다른 종류로 덮을 때 색이 낡은 채로 남는다.
    head_rev = (len(b.disks), hash(tuple(b.disk_kinds)), bool(es._flowed),
                hash(frozenset(es._wet_keys)) if es._wet_keys else 0)
    head_fresh = full or sess.get("head_rev") != head_rev
    if head_fresh:
        sess["head_rev"] = head_rev
    else:
        keep.append("heads")

    # 물길은 손질이 망을 건드리면 지워지고, 물흐름을 돌려야 다시 생긴다.
    water = sess.get("water_path") or []
    wet_rev = (len(water), rev)
    wet_fresh = full or sess.get("wet_rev") != wet_rev
    if wet_fresh:
        sess["wet_rev"] = wet_rev
    else:
        keep.append("wet_pipes")

    # 후보 점선은 훑기·지우기·적용 때만 바뀐다. 상한(4,000곳)까지 차면 응답마다
    # 190KB 라, 클릭할 때마다 다시 실어 보낼 이유가 없다.
    aj_fresh = full or sess.get("aj_sent") != sess.get("aj_seq", 0)
    if aj_fresh:
        sess["aj_sent"] = sess.get("aj_seq", 0)
    else:
        keep.append("autojoin")

    # 최불리 경로는 간선 697개(실측)라 30KB 다. 나머지를 1KB 로 줄여놓고 이것만
    # 매번 실어 보낼 이유가 없다. 지문은 내용에서 뽑는다 — 따로 세어 둘 것이 없다.
    w = sess.get("worst")
    worst_rev = (None if not w else
                 (len(w["heads"]), w["far_m"], w["near_m"],
                  len(w["edges"]), w.get("sheet")))
    worst_fresh = full or sess.get("worst_rev") != worst_rev
    if worst_fresh:
        sess["worst_rev"] = worst_rev
    else:
        keep.append("worst")

    kinds: dict[str, int] = {}
    for k in b.disk_kinds:
        kinds[k] = kinds.get(k, 0) + 1
    return {
        "mode": es.mode,
        "counts": {"pts": len(b.pts), "edges": len(b.edges),
                   "heads": len(b.disks),
                   "bodies": (len(g["body_groups"]) if net_fresh
                              else sess.get("n_bodies", 0)),
                   "joins": len(b.joins), "deletes": len(b.deletes)},
        "kinds": kinds,
        # 화면에만 쓰는 좌표다 — 도면 한 장이 수백 m 인데 0.1mm 를 실어 나를
        # 이유가 없다. 평평한 배열(찍기 캔버스가 이미 쓰는 규약) + 정수 mm.
        "body_groups": [
            {"css": color,
             "segs": [v for a, c in segs
                      for v in (round(a[0]), round(a[1]),
                                round(c[0]), round(c[1]))]}
            for segs, color in g["body_groups"]],
        "heads": ([] if not head_fresh else
                  [[_r1(d[0]), _r1(d[1]), _r1(d[2]), css]
                   for d, css in g["heads"]]),
        "multi_heads": ([] if not head_fresh else
                        [[_r1(d[0]), _r1(d[1]), _r1(d[2])]
                         for d in g["multi_heads"]]),
        "sources": [[_r1(x), _r1(y)] for x, y in g["sources"]],
        "valves": [[_r1(x), _r1(y)] for x, y in g["valves"]],
        "pending": ([[_r1(g["pending"][0][0]), _r1(g["pending"][0][1])],
                     [_r1(g["pending"][1][0]), _r1(g["pending"][1][1])]]
                    if g["pending"] else None),
        "selected_head": ([_r1(v) for v in g["selected_head"]]
                          if g["selected_head"] else None),
        # E 의 wet_pipes 는 연출 프레임 전용이라 연출이 끝나면 빈다(헤드 색으로
        # 결과가 남는 구조). 브라우저는 연출을 돌리지 않으므로, 물이 닿은 간선
        # 전체를 따로 붙잡아 두고 그 위에 겹쳐 그린다. 손질을 건드리면 지운다.
        "wet_pipes": ([] if not wet_fresh else
                      [[_r1(a[0]), _r1(a[1]), _r1(c[0]), _r1(c[1])]
                       for a, c in g["wet_pipes"]] or water),
        "wet_counts": es.wet_kind_counts(),
        "flowed": bool(es._flowed),
        "worst": _worst_view(sess) if worst_fresh else None,
        # 아직 «후보» 다 — 캔버스에서 점선·다른 색으로만 그린다. 실측 연결선과
        # 섞어 그리면 사람이 확인한 것과 기계가 고른 것을 구별할 수 없다.
        "autojoin": _autojoin_view(sess) if aj_fresh else None,
        # 잡 응답에는 결과가 실리지 않는다(`_job_view` 는 진행 줄만 보낸다).
        # 방금 붙인 결과는 손질 상태에 실어 화면이 한 곳에서 읽게 한다.
        "autojoin_report": sess.get("autojoin_report"),
        "sheets": sess.get("sheets") or [],
        "body_stat": sess.get("body_stat") or _body_stat(b),
        "palette": {"source": EDIT_SOURCE, "valve": EDIT_VALVE,
                    "wet": EDIT_WET_PIPE, "kinds": dict(KIND_COLORS)},
        # [F-10d · D-F10-5] 마지막 최불리 계산 뒤로 몇 건을 고쳤나. 수정마다
        # 다시 돌리지 않는 대신(검출 실측 ~18초) 이 수를 화면에 배지로 띄우고,
        # 「다시 계산」은 사람이 한 번만 누른다.
        "edits_since_worst": int(sess.get("worst_edits") or 0),
        # 이 이름들은 «안 바뀌었으니 들고 있던 것을 그대로 쓰라» 는 뜻이다.
        # 빈 배열만으로는 «비었다» 와 구별되지 않아 물길이 조용히 사라졌었다.
        "keep": keep,
        "bounds": _pts_bounds(b.pts),
    }


def _autojoin_view(sess: dict) -> dict | None:
    """자동 이음 후보 — 화면에는 점선으로만 나간다."""
    sc = sess.get("autojoin")
    if not sc:
        return None
    return {
        "eps_mm": sc["eps_mm"], "auto_eps_mm": sc["auto_eps_mm"],
        "n": len(sc["cands"]), "ends": sc["ends"], "near": sc["near"],
        "kept": sc["kept"], "by_kind": sc["by_kind"],
        "dropped": sc["dropped"], "bodies_before": sc["bodies_before"],
        "trials": sc["trials"],
        "lines": [c["line"] for c in sc["cands"]],
    }
