# -*- coding: utf-8 -*-
"""모듈 A 에서 빌려온 것 — 최불리 K · 도면 장 나누기 · 범위 제한 · PIPENET."""
from __future__ import annotations

import math

from routes.module_f.common import REMOTE_K_DEFAULT, _r1


def _worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT,
                   only_heads=None, source_index=None) -> dict:
    """[F-0·D1] 엔진(G design/worst.py)으로 위임 — 구현은 한 벌만 둔다.

    이 파일에 있던 원본 구현이 G1 때 엔진으로 옮겨 갔고, 여기 남아 있던
    복제본은 F-1(급수원 지정)부터 어긋날 판이었다. 같은 알고리즘이 두 벌이면
    한쪽만 고쳐지는 날이 반드시 온다 — 껍데기만 남기고 엔진을 부른다.
    (import 는 지연 — _boot() 가 sys.path 에 엔진을 올린 뒤라야 열린다.)
    """
    from services.cad_import.design.worst import worst_k_heads
    return worst_k_heads(pts, edges, hnodes, sources, k=k,
                         only_heads=only_heads, source_index=source_index)


def _worst_view(sess: dict) -> dict | None:
    """화면용 — 최불리 배관망(corridor)·앵커·최원 유하거리 경로·담당 헤드 수.

    corridor 간선은 좌표 4개 + load 를 함께 싣는다(화면이 굵기/색을 정한다).
    앵커는 «가장 불리한 지점» 이라 따로 강조한다.

    ★`anchor_path` 는 corridor 와 **겹치는 부분집합**이다. 그래도 따로 싣는다 —
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
    an = w.get("anchor")
    path = [n for n in (w.get("anchor_path") or ()) if 0 <= n < len(pts)]
    return {
        "k": len(w["heads"]),
        "reachable": w["reachable"],
        "far_m": w["far_m"],
        "near_m": w["near_m"],
        "span_m": w.get("span_m", 0.0),
        "total_m": w.get("total_m", 0.0),
        "max_load": w.get("max_load", 0),
        "sheet": w.get("sheet"),
        # [F-1] 어느 급수원 기준의 최불리인지 — 화면이 이것을 그대로 보여 준다.
        "source": w.get("source_tag"),
        # 사람이 가둔 영역 — 다시 그릴 수 있게 그대로 돌려준다.
        "zones": [[_r1(v) for v in z] for z in (w.get("zones") or ())],
        "candidates": w.get("candidates", w["reachable"]),
        "heads": [[_r1(disks[hi][0]), _r1(disks[hi][1]), _r1(disks[hi][2])]
                  for hi in w["heads"] if hi < len(disks)],
        "anchor": ([_r1(disks[an][0]), _r1(disks[an][1]), _r1(disks[an][2])]
                   if isinstance(an, int) and an < len(disks) else None),
        # 급수원 → 앵커. 절점 열을 그대로 준다(화면이 한 줄로 잇는다).
        "anchor_path": [[_r1(pts[n][0]), _r1(pts[n][1])] for n in path],
        "anchor_path_m": w.get("anchor_path_m", 0.0),
        "corridor": [[_r1(pts[a][0]), _r1(pts[a][1]),
                      _r1(pts[c][0]), _r1(pts[c][1]), int(load)]
                     for (a, c), load in w.get("loads", {}).items()],
    }


# ────────────────────────────────────────────── 자동 이음 · 도면 장 · 덩이


def _sheet_frames(board) -> list[dict]:
    """한 파일에 도면이 여러 장 들어 있는지 — 모듈 A 의 규칙을 그대로 부른다.

    국내 도서는 도면 한 장이 곧 파일 하나가 아니다(A 실측 — 죽전 6장·청라
    포레스트 3장·대구오페라 단위세대 5장). 여러 장을 한 망으로 보면 최불리 30 이
    서로 다른 도면의 헤드를 섞어 뽑아 계산이 성립하지 않는다.

    A 의 `detect_sheet_frames` 는 헤드 좌표(`.pos`)만 본다 — 문턱도 상수가 아니라
    그 도면의 헤드 간격에서 잰다. 그래서 규칙을 베끼지 않고 그대로 호출한다.
    """
    disks = getattr(board, "disks", None) or ()
    if len(disks) < 24:
        return []

    class _Head:  # A 가 보는 것은 .pos 하나뿐이다
        __slots__ = ("pos",)

        def __init__(self, p):
            self.pos = p

    try:
        from remote30_prototype import detect_sheet_frames
    except Exception as exc:  # noqa: BLE001 — A 가 없어도 손질은 돌아야 한다
        print(f"[손질] 도면 장 나누기 건너뜀 — 모듈 A 미탑재: {exc}")
        return []
    try:
        return detect_sheet_frames(
            [_Head((float(d[0]), float(d[1]))) for d in disks])
    except Exception as exc:  # noqa: BLE001
        print(f"[손질] 도면 장 나누기 실패: {exc}")
        return []


def _restrict_to_worst(payload: dict, board, worst: dict) -> dict:
    """변환 대상을 최불리 K 헤드로 좁힌다 — 헤드만 지우고 배관은 안 자른다.

    간선을 직접 잘라내고 싶은 유혹이 있지만 그러면 안 된다. 모듈 E 의
    `build_planar_graph` 는 이미 «급수원에서 물 닿는 간선만 남기고, 헤드로
    가지 않는 막다른관을 쳐내는» 단계를 갖고 있다(실측 로그: 물길 필터 →
    막다른관 삭제). 그러니 남길 헤드만 남겨 두면 그 배관은 E 가 제 규칙으로
    정리한다. 손으로 자르면 E 가 지키는 불변식(티 겹침·노드정리)을 깬다.

    hcov / disk_kinds / head_kinds 는 같은 디스크 집합을 가리키므로 함께 건다.
    ups 는 좌표 집합으로만 쓰여 남아 있어도 해가 없다.
    """
    from services.cad_import.kinds import disk_key

    keep_idx = {int(i) for i in (worst or {}).get("heads") or ()}
    disks = list(board.disks)
    kept = [disks[i] for i in sorted(keep_idx) if 0 <= i < len(disks)]
    if not kept:
        return payload

    keys = {disk_key(d[0], d[1], d[2]) for d in kept}
    out = dict(payload)
    out["hcov"] = [list(d) for d in kept]
    dk = payload.get("disk_kinds") or []
    out["disk_kinds"] = [dk[i] for i in sorted(keep_idx) if 0 <= i < len(dk)]

    fresh = []
    for rec in payload.get("head_kinds") or ():
        if not isinstance(rec, dict) or "c" not in rec:
            continue
        c = rec["c"]
        r = rec.get("head_r")
        if r is None and rec.get("tri_side"):
            r = float(rec["tri_side"]) / math.sqrt(3.0)
        if r is None:
            continue
        if disk_key(c[0], c[1], r) in keys:
            fresh.append(dict(rec))
    out["head_kinds"] = fresh
    print(f"[변환] 최불리 {len(kept)} 헤드로 범위를 좁힘 "
          f"(도면 헤드 {len(disks)} · 종류표 {len(fresh)}행)")
    return out


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
