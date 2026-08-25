# -*- coding: utf-8 -*-
"""KFP 변환(직렬화) 시점 전용 — 티 허브 겹침 정규화.

    from services.cad_import.convert.normalize import normalize_tee_overlaps

2단계 이음 로직(_tmp_flow_water)은 건드리지 않는다 [2026-08-11 오너 확정].
임포트 이음이 만든 «관통 배관 + 티 허브 스텁» 겹침만 변환 때 바로잡는다.

확인된 패턴 (MF101_흰색점선범위_유저정리3.kfp 실측 · 겹침 35쌍 전부 이 모양):

    u ──────────────────── v   관통 간선 (티 허브 위를 지나감)
    u ── h                     스텁 간선 (~0.15m · 관통과 완전 겹침)
         │
         가지(branch)          h = 티 허브 (가지가 달린 노드)

목표 모양:  u ── h ── v  (main—tee—main · 가지는 h 에 그대로)
              │
            가지

고치는 방법: 관통 u—v 삭제, h—v 추가. 스텁 u—h 는 그대로 둔다(그것이
main 의 앞쪽 반이다). 새 선을 «만들지» 않는다 — 있던 기하를 허브에서
쪼갤 뿐이라 총 기하·물길은 그대로다.

패턴이 정확히 맞을 때만 고친다 (정석 아니면 그대로):
  · h 가 u—v 선분 «안»(양끝 제외)에 횡이탈 0으로 정확히 얹혀 있고
  · h 에 스텁 말고 다른 간선이 하나 이상(가지) 달려 있을 때만.
    매달린 점(스텁뿐 · 차수 1)은 확인된 패턴이 아니므로 건너뛴다.

배관 속성: 이 변환은 모든 배관을 같은 속성(DN25·C120·KSD3507)으로 만들므로
쪼갠 두 토막도 자동으로 관통과 같은 속성이다.
"""
from __future__ import annotations

import math
import unittest


def normalize_tee_overlaps(pos, edges, eps=1e-6):
    """관통 간선을 티 허브에서 쪼개 겹침을 없앤다.

    pos   : {vid: (x, y)} — 같은 눈금(격자 스냅 뒤) 좌표
    edges : {(a, b), …} 무방향 간선 (a != b)
    eps   : «정확히 얹힘» 판정 여유. 스냅 좌표라 사실상 0 비교다.

    반환: (정규화된 간선 frozenset, 쪼갠 관통 간선 수)
    """
    cur = {tuple(sorted(e)) for e in edges}
    n_split = 0
    for _ in range(1000):
        adj = {}
        for a, b in cur:
            adj.setdefault(a, set()).add(b)
            adj.setdefault(b, set()).add(a)

        def hub_inside(u, v):
            """u 의 이웃 허브 중 u—v 선분 안에 정확히 얹힌 것.

            여럿이면 u 에서 «먼» 것 — 남는 겹침(u 쪽 스텁끼리)은
            다음 라운드가 같은 규칙으로 잡는다.
            """
            ux, uy = pos[u]
            vx, vy = pos[v]
            dx, dy = vx - ux, vy - uy
            length = math.hypot(dx, dy)
            if length <= eps:
                return None
            best = None
            for h in adj.get(u, ()):
                if h == v or len(adj.get(h, ())) < 2:
                    continue
                wx, wy = pos[h][0] - ux, pos[h][1] - uy
                along = (wx * dx + wy * dy) / length
                if not (eps < along < length - eps):
                    continue
                lat2 = (wx * wx + wy * wy) - along * along
                if lat2 > eps * eps:
                    continue
                if best is None or along > best[0]:
                    best = (along, h)
            return None if best is None else best[1]

        changed = False
        for (u, v) in sorted(cur, key=str):
            if (u, v) not in cur:
                continue
            h = hub_inside(u, v)
            if h is None:
                h = hub_inside(v, u)
                if h is None:
                    continue
                u, v = v, u          # 스텁이 v 쪽 — 대칭으로 처리
            cur.discard(tuple(sorted((u, v))))
            cur.add(tuple(sorted((h, v))))
            n_split += 1
            changed = True
        if not changed:
            break
    return frozenset(cur), n_split


class TeeNormalizeTest(unittest.TestCase):
    def test_confirmed_pattern_split(self):
        # 실측 N677 모양: 관통 v—u 가 허브 h(u에서 0.15) 위를 지나감
        pos = {"u": (43.6, 19.05), "v": (42.0, 19.05),
               "h": (43.45, 19.05), "b1": (43.45, 18.5), "b2": (43.45, 21.0)}
        edges = {("u", "v"), ("u", "h"), ("h", "b1"), ("h", "b2")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(1, n)
        self.assertEqual(
            {("h", "u"), ("h", "v"), ("b1", "h"), ("b2", "h")}, set(out))

    def test_stub_on_v_side(self):
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0),
               "h": (1.85, 0.0), "b": (1.85, 1.0)}
        edges = {("u", "v"), ("v", "h"), ("h", "b")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(1, n)
        self.assertEqual({("h", "u"), ("h", "v"), ("b", "h")}, set(out))

    def test_dangling_stub_untouched(self):
        # 허브에 가지가 없다(차수 1) → 확인된 패턴 아님, 그대로
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0), "h": (0.15, 0.0)}
        edges = {("u", "v"), ("u", "h")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(0, n)
        self.assertEqual({("u", "v"), ("h", "u")}, set(out))

    def test_offline_hub_untouched(self):
        # 허브가 관통 선 위가 아님(횡이탈 0.05) → 그대로
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0),
               "h": (0.15, 0.05), "b": (0.15, 1.0)}
        edges = {("u", "v"), ("u", "h"), ("h", "b")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(0, n)

    def test_hub_beyond_end_untouched(self):
        # 허브가 관통 구간 밖(뒤쪽) → 겹침 없음, 그대로
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0),
               "h": (-0.15, 0.0), "b": (-0.15, 1.0)}
        edges = {("u", "v"), ("u", "h"), ("h", "b")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(0, n)

    def test_two_hub_chain(self):
        # 한 끝에 허브 둘(0.15 · 0.3) → 먼 것부터 쪼개고 다음 라운드가 마무리
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0),
               "h1": (0.15, 0.0), "b1": (0.15, 1.0),
               "h2": (0.3, 0.0), "b2": (0.3, 1.0)}
        edges = {("u", "v"), ("u", "h1"), ("u", "h2"),
                 ("h1", "b1"), ("h2", "b2")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(2, n)
        self.assertEqual(
            {("h1", "u"), ("h1", "h2"), ("h2", "v"),
             ("b1", "h1"), ("b2", "h2")}, set(out))

    def test_existing_half_deduped(self):
        # h—v 가 이미 있으면(삼각) 관통 삭제만 — 간선 수가 준다
        pos = {"u": (0.0, 0.0), "v": (2.0, 0.0),
               "h": (0.15, 0.0), "b": (0.15, 1.0)}
        edges = {("u", "v"), ("u", "h"), ("h", "v"), ("h", "b")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(1, n)
        self.assertEqual({("h", "u"), ("h", "v"), ("b", "h")}, set(out))

    def test_clean_net_untouched(self):
        # 이미 정상인 티(main—tee—main) → 무변경
        pos = {"u": (0.0, 0.0), "h": (1.0, 0.0), "v": (2.0, 0.0),
               "b": (1.0, 1.0)}
        edges = {("u", "h"), ("h", "v"), ("h", "b")}
        out, n = normalize_tee_overlaps(pos, edges)
        self.assertEqual(0, n)
        self.assertEqual({tuple(sorted(e)) for e in edges}, set(out))


if __name__ == "__main__":
    unittest.main(verbosity=2)
