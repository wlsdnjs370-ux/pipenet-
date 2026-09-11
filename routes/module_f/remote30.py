# -*- coding: utf-8 -*-
"""모듈 A 에서 빌려온 것 — 최불리 K · 도면 장 나누기 · 범위 제한 · PIPENET."""
from __future__ import annotations

import math

from routes.module_f.common import REMOTE_K_DEFAULT, _r1


def _worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT,
                   only_heads=None, source_index=None, head_xy=None) -> dict:
    """[F-0·D1] 엔진(G design/worst.py)으로 위임 — 구현은 한 벌만 둔다.

    이 파일에 있던 원본 구현이 G1 때 엔진으로 옮겨 갔고, 여기 남아 있던
    복제본은 F-1(급수원 지정)부터 어긋날 판이었다. 같은 알고리즘이 두 벌이면
    한쪽만 고쳐지는 날이 반드시 온다 — 껍데기만 남기고 엔진을 부른다.
    (import 는 지연 — _boot() 가 sys.path 에 엔진을 올린 뒤라야 열린다.)
    """
    from services.cad_import.design.worst import worst_k_heads
    return worst_k_heads(pts, edges, hnodes, sources, k=k,
                         only_heads=only_heads, source_index=source_index,
                         head_xy=head_xy)


def _design_corridor(sess: dict) -> dict | None:
    """[배관망도 한 벌로] 표가 있으면 **표의 배관망**을 board mm 로 돌려준다.

    ■ 왜 필요한가 (2026-09-11 사용자 지적 · 대명동 K=30 · 영역 2곳)

      「최불리 배관망에서 등록했던 배관망을 그대로 아이소에 가져오는 것조차
       안 된다.」 선정(§2-1)은 맞췄는데 **배관망이 달랐다.** 실측:

          corridor 461선분 156.1 m   /   표 202선분 160.6 m
          표에만 있는 긴 배관 3.35 · 3.00 · 2.96 · 2.34 · 2.03 m …  (합 11.1 m)

      ★막다른 관이 남아서가 아니다. 표의 막다른 끝은 5곳뿐이고(0.03~0.50 m)
        전부 정상이다 — 2곳은 급수원 스텁, 3곳은 「끝배관 1단 보호」다
        (전개 로그의 `끝배관1단보호 3` 과 정확히 맞는다).

      원인은 **두 그림이 서로 다른 그래프에서 최단경로를 뽑는 것**이다:
        corridor 는 손질 board(간선 2,357) 위에서,
        표는 물길필터·티 겹침 정규화·직선 복원·노드정리를 거친 망 위에서.
      고리가 있는 망에서는 길이가 비슷한 **다른 길**이 뽑힌다. 둘 다 «맞는»
      길이지만 화면 둘이 다른 그림을 내놓는다.

    ■ 규칙

      표가 서 있고 그 표가 **지금 선정·지금 손질판의 것**이면(옛 표가 아니면),
      평면 보기는 board 최단경로가 아니라 **표의 배관망**을 그린다. 그러면
      평면과 아이소가 «같은 망» 이 된다 — 구성으로 보장된다.
      표가 없거나 옛 것이면 종전대로 board corridor 를 그린다.

    반환: {"corridor", "path", "max_load", "total_m"} 또는 None.
    """
    d = sess.get("design") or {}
    tbl, got = d.get("tables"), d.get("got")
    if not tbl or not got:
        return None
    # [리팩터링 2026-09-11] 판단은 selection.py 에 산다 — 라우트 파일
    #   (api_design)을 import 하던 것을 끊었다. 순수 판단이 라우트에 살면
    #   이런 의존이 늘고, 언젠가 순환 import 로 돌아온다.
    from routes.module_f.selection import _design_stale, _load_map
    if _design_stale(sess):
        # 옛 표를 «지금 망» 으로 그리면 안 된다 — 그것이 바로 고치려던 증상이다.
        return None
    origin = got.get("origin_mm")
    if not origin:
        return None
    ox = float(origin[0]) - 1000.0
    oy = float(origin[1]) - 1000.0
    at = {str(n.get("label")): (float(n.get("x") or 0) + ox,
                                float(n.get("y") or 0) + oy)
          for n in (getattr(tbl, "nodes", None) or ())}
    lm = _load_map(got)
    segs: list = []
    adj: dict = {}
    for pr in (getattr(tbl, "pipes", None) or ()):
        a, c = str(pr.get("in")), str(pr.get("out"))
        pa, pc = at.get(a), at.get(c)
        adj.setdefault(a, []).append(c)
        adj.setdefault(c, []).append(a)
        if not pa or not pc or math.dist(pa, pc) <= 1.0:
            continue          # 세로 구간 — 평면에서는 점 하나다
        segs.append([_r1(pa[0]), _r1(pa[1]), _r1(pc[0]), _r1(pc[1]),
                     int(lm.get(str(pr.get("label")), 0) or 0)])
    if not segs:
        return None

    # 최원 유하거리 «경로» — 급수원(Input) → 기준 헤드. 표 위에서 다시 잇는다.
    #   board 쪽 경로를 그대로 쓰면 그린 망과 다른 줄이 되어 더 헷갈린다.
    root = next((str(n.get("label")) for n in (getattr(tbl, "nodes", None) or ())
                 if str(n.get("io_node") or "").strip().lower()
                 not in ("", "no", "none")), None)
    w = sess.get("worst") or {}
    anchor = None
    try:
        hi = int(w.get("worst_head"))
        disks = sess["edit"].board.disks
        if 0 <= hi < len(disks):
            q = (float(disks[hi][0]), float(disks[hi][1]))
            noz = {str(z.get("in")) for z in (getattr(tbl, "nozzles", None) or ())}
            cand = [(math.dist(p, q), lab) for lab, p in at.items() if lab in noz]
            if cand:
                anchor = min(cand)[1]
    except (TypeError, ValueError, KeyError, AttributeError):
        anchor = None
    path: list = []
    if root and anchor:
        prev = {root: None}
        queue = [root]
        while queue:
            cur = queue.pop(0)
            if cur == anchor:
                break
            for nxt in adj.get(cur, ()):
                if nxt not in prev:
                    prev[nxt] = cur
                    queue.append(nxt)
        if anchor in prev:
            cur, chain = anchor, []
            while cur is not None:
                chain.append(cur)
                cur = prev[cur]
            chain.reverse()
            path = [[_r1(at[n][0]), _r1(at[n][1])] for n in chain if n in at]
    total = sum(math.dist((s[0], s[1]), (s[2], s[3])) for s in segs) / 1000.0
    return {"corridor": segs, "path": path,
            "max_load": max((s[4] for s in segs), default=0),
            "total_m": round(total, 2)}


def _worst_view(sess: dict) -> dict | None:
    """화면용 — 최불리 배관망(corridor)·앵커·최원 유하거리 경로·담당 헤드 수.

    corridor 간선은 좌표 4개 + load 를 함께 싣는다(화면이 굵기/색을 정한다).
    앵커는 «가장 불리한 지점» 이라 따로 강조한다.

    ★`worst_path` 는 corridor 와 **겹치는 부분집합**이다. 그래도 따로 싣는다 —
      far_m 이 곧 이 줄의 길이인데, corridor 를 굵기로만 그리면 그 거리가 어느
      줄인지 화면에서 읽을 수 없다. 기준압을 잡는 지점이 앵커라면 그 압이 어느
      관을 타고 오는지도 같이 보여야 관경을 키울지 경로를 줄일지 정할 수 있다.
    """
    w = sess.get("worst")
    if not w:
        return None
    b = sess["edit"].board
    pts = b.pts
    disks = b.disks
    an = w.get("worst_head")
    path = [n for n in (w.get("worst_path") or ()) if 0 <= n < len(pts)]
    # ★[배관망도 한 벌로] 표가 서 있고 그것이 «지금» 것이면, 평면 보기는
    #   board 최단경로가 아니라 **표의 배관망**을 그린다. 실측(대명동 K30 ·
    #   영역 2곳): 두 그림이 서로 다른 그래프에서 최단경로를 뽑아 11.1 m 가
    #   달랐다 — 고리가 있는 망에서 길이가 비슷한 다른 길을 고른 것이다.
    dc = _design_corridor(sess)
    return {
        "k": len(w["heads"]),
        "reachable": w["reachable"],
        "far_m": w["far_m"],
        "near_m": w["near_m"],
        "span_m": w.get("span_m", 0.0),
        # 설계면적 직사각형의 실제 크기 — 화면이 «몇 ㎡ 인가» 를 말한다.
        "area_w_m": w.get("area_w_m", 0.0),
        "area_h_m": w.get("area_h_m", 0.0),
        "area_m2": w.get("area_m2", 0.0),
        "total_m": (dc["total_m"] if dc else w.get("total_m", 0.0)),
        "max_load": (dc["max_load"] if dc else w.get("max_load", 0)),
        # 지금 그리는 망이 «표의 것» 인가 «손질 최단경로» 인가 — 화면이 말한다.
        "net_from": ("design" if dc else "edit"),
        "sheet": w.get("sheet"),
        # [F-1] 어느 급수원 기준의 최불리인지 — 화면이 이것을 그대로 보여 준다.
        "source": w.get("source_tag"),
        # ★[두 화면 선정일치 §2-1] 지금 그리는 것이 «표에 들어간 선정» 인가,
        #   아직 «손질이 고른 것» 인가. 표를 확정하면 선정이 교체될 수 있으므로
        #   (못 붙는 헤드를 다음 순위로 채운다) 화면이 어느 쪽을 보고 있는지
        #   말해야 한다 — 같은 그림에 두 뜻이 있으면 사람이 판단을 못 한다.
        "from_design": bool(w.get("from_design")),
        # 사람이 가둔 영역 — 다시 그릴 수 있게 그대로 돌려준다.
        "zones": [[_r1(v) for v in z] for z in (w.get("zones") or ())],
        "candidates": w.get("candidates", w["reachable"]),
        "heads": [[_r1(disks[hi][0]), _r1(disks[hi][1]), _r1(disks[hi][2])]
                  for hi in w["heads"] if hi < len(disks)],
        "worst_head": ([_r1(disks[an][0]), _r1(disks[an][1]), _r1(disks[an][2])]
                   if isinstance(an, int) and an < len(disks) else None),
        # 급수원 → 앵커. 절점 열을 그대로 준다(화면이 한 줄로 잇는다).
        #   표가 있으면 그 줄도 **표 위에서** 다시 잇는다 — 그린 망과 다른 줄을
        #   덧그리면 더 헷갈린다.
        "worst_path": (dc["path"] if dc and dc["path"] else
                       [[_r1(pts[n][0]), _r1(pts[n][1])] for n in path]),
        "worst_path_m": w.get("worst_path_m", 0.0),
        "corridor": (dc["corridor"] if dc else
                     [[_r1(pts[a][0]), _r1(pts[a][1]),
                       _r1(pts[c][0]), _r1(pts[c][1]), int(load)]
                      for (a, c), load in w.get("loads", {}).items()]),
    }


# ────────────────────────────────────────────── 자동 이음 · 도면 장 · 덩이


def _sheet_frames(board) -> list[dict]:
    """[D1] 엔진(G design/worst.py)으로 위임 — 구현은 한 벌만 둔다.

    같은 몸통이 여기와 엔진에 두 벌 있었다(2026-09-11 diff 로 확인 — 차이는
    로그 접두어와 엔진 쪽 sys.path 보강뿐). 두 벌이면 한쪽만 고쳐지는 날이
    반드시 온다 — `_worst_k_heads` 와 같은 결정이다. 로그 접두어는 엔진 것
    ([손질]→[G])을 따른다. (import 는 지연 — _boot 뒤라야 엔진이 열린다.)
    """
    from services.cad_import.design.worst import sheet_frames
    return sheet_frames(board)


def _restrict_to_worst(payload: dict, board, worst: dict) -> dict:
    """[D1] 엔진(G design/restrict.py)으로 위임 — 구현은 한 벌만 둔다.

    같은 몸통이 여기와 엔진에 두 벌 있었다(2026-09-11 diff 로 확인 — 차이는
    로그 접두어 [변환]/[G2] 뿐). 설계 경로는 이미 엔진 판을 쓰고 있었으니,
    변환 경로(api_convert)만 이 사본을 물고 있었다 — 한쪽만 고쳐지는 날이
    오기 전에 합친다. 로그 접두어는 엔진 것([G2])을 따른다.
    """
    from services.cad_import.design.restrict import restrict_to_worst
    return restrict_to_worst(payload, board, worst)


# [정리 2026-08-31] `_emit_pipenet(sess, kfp, out_dir)` 를 지웠다.
#
#   도크스트링이 「진단 스크립트 호환으로만 남긴다」고 적혀 있었는데 **그 이유가
#   더는 참이 아니다** — 저장소 전체(비추적 파일 포함)를 훑어도 부르는 곳이 없다.
#   남은 언급은 리팩터링 이전 스냅샷(`data/_module_f_before_refactor.py`)과
#   작업지시서뿐이다. 틀린 이유를 단 채로 코드를 두면 다음 사람이 그 이유를
#   믿고 손대지 못한다 — 죽은 코드보다 «죽은 이유» 가 더 오래 남는다.
#
#   기능 자체는 D3 로 은퇴했다 — 설계구역 없는 전체망 SDF 는 수리계산 입력이
#   아니다. 수리계산 입력 SDF 는 `design/emit`(G 엔진)이 만든다.
