# -*- coding: utf-8 -*-
"""모듈 D — D4 라벨 배치기.

배치기 자체의 계약만 본다. "실물 도면에서 겹침 0" 은 만든 PDF 에서 글자를 도로
뽑아 재는 검사라 ``test_iso_renderer.py`` 가 맡는다 — 배치기가 쓴 상자를 돌려 쓰면
상자를 잘못 재도 양쪽이 똑같이 잘못 재서 통과해 버린다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_label_layout import (  # noqa: E402
    LEADER_AFTER,
    LabelRequest,
    box_of,
    lay_out,
    overlapping_pairs,
)

# 글자 하나가 가로 1, 한 줄 높이 1 인 자. 상자 계산만 보므로 글꼴은 끌어들이지 않는다.
def measure(text: str) -> tuple[float, float]:
    return float(len(text)), 1.0


def _requests(n: int, *, step: float = 0.0, kind: str = "link", angle: float = 0.0):
    return [LabelRequest(f"L{i}", "1234", (i * step, 0.0), angle, kind) for i in range(n)]


def test_crowded_labels_are_pushed_apart_not_stacked():
    # 같은 자리에 몰아넣어도 겹친 채 두지 않는다.
    placed, report = lay_out(_requests(40), measure)
    assert report.placed == 40 and report.dropped == ()
    assert overlapping_pairs(placed, measure) == []


def test_obstacles_are_avoided():
    # 도면 주기처럼 못 움직이는 글자 위에는 놓지 않는다.
    wall = (-6.0, -3.0, 6.0, 3.0)
    placed, report = lay_out(_requests(12), measure, obstacles=[wall])
    assert report.placed == 12
    for lab in placed:
        w, h = box_of(measure, lab.text, lab.angle)
        assert not (lab.x - w / 2 < wall[2] and wall[0] < lab.x + w / 2
                    and lab.y - h / 2 < wall[3] and wall[1] < lab.y + h / 2)


def test_far_pushed_labels_get_a_leader_back_to_the_anchor():
    placed, report = lay_out(_requests(40), measure)
    assert report.leaders == sum(1 for lab in placed if lab.leader)
    assert 0 < report.leaders < len(placed)
    for lab in placed:
        _, h = box_of(measure, lab.text, lab.angle)
        reach = max(abs(lab.x), abs(lab.y))
        if lab.leader is None:
            assert reach < h * (LEADER_AFTER + 1.0)
        else:
            # 지시선은 기준점에서 출발해 상자 모서리에서 끝난다.
            assert lab.leader[0] == (0.0, 0.0)
            assert reach > h * LEADER_AFTER


def test_link_labels_start_perpendicular_to_the_pipe():
    # 관로 라벨은 관을 덮지 않도록 관 직각 방향부터 본다.
    placed, _ = lay_out([LabelRequest("A", "1234", (0.0, 0.0), 0.0, "link")], measure)
    assert abs(placed[0].x) < 1e-9 and abs(placed[0].y) > 0


def test_nozzle_labels_hang_below_the_discharge():
    # 노즐 위쪽은 노즐이 매달린 가지배관 자리다.
    placed, _ = lay_out([LabelRequest("A", "1234", (0.0, 0.0), 0.0, "nozzle")], measure)
    assert abs(placed[0].x) < 1e-9 and placed[0].y < 0


def test_labels_with_no_room_are_reported_not_overlapped():
    # 후보 자리를 다 훑고도 빈 곳이 없으면 겹쳐 그리지 않고 리포트에 싣는다.
    placed, report = lay_out(_requests(600), measure)
    assert report.dropped
    assert report.placed + len(report.dropped) == 600
    assert overlapping_pairs(placed, measure) == []


def test_budget_stops_the_layout_and_names_what_was_left_out():
    placed, report = lay_out(_requests(400), measure, budget_s=0.0)
    assert report.dropped and report.seconds < 3.0
    assert len(placed) + len(report.dropped) == 400


def test_layout_is_deterministic():
    # 한 라벨이 움직여 온 도면이 재배치되면 시각 회귀 비교가 성립하지 않는다.
    first, _ = lay_out(_requests(60, step=0.7), measure)
    second, _ = lay_out(_requests(60, step=0.7), measure)
    assert [(p.key, p.x, p.y) for p in first] == [(p.key, p.x, p.y) for p in second]


def test_rotated_box_covers_the_rotated_text():
    flat = box_of(measure, "12345678", 0.0)
    assert box_of(measure, "12345678", 90.0) == pytest.approx(flat[::-1], abs=1e-9)
    # 비스듬한 글자는 어느 축으로도 눕힌 글자보다 두껍다.
    slanted = box_of(measure, "12345678", 45.0)
    assert slanted[0] > flat[1] and slanted[1] > flat[1]


def test_empty_input_reports_nothing_placed():
    placed, report = lay_out([], measure)
    assert placed == [] and report.placed == 0 and report.dropped == ()
