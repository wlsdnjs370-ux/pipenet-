# -*- coding: utf-8 -*-
"""[변환기] 배관장 권위 — 좌표는 «선언과 이미 일치할 때만» 믿는다.

■ 실측 사고 (2026-09 · 사용자 「그외에 다른 로직 누수도 조치」에서 발굴)

  `parse_sdf` 는 「length/좌표거리 비의 중앙값이 1 근처면 좌표거리로 length 를
  덮는다」였다. 모듈 F 결합 SDF 는 스키매틱 좌표(비 0.02~200 제각각)인데
  중앙값만 우연히 1 근처에 떨어져 — **선언 길이 58개가 전부 좌표거리로
  갈아치워졌다**(9.735 m → 0.047 m · 총연장 183.5 → 155.9 m). 그 값이
  KFP·HAS 수리계산 입력으로 그대로 나갔고, 세 형식이 «같은 오염원» 에서
  나와 서로 일치했으므로 cross_check(S760)도 잡지 못했다.

  화면에서는 「계통도·상하향식 길이가 쪼개져 깨져 보인다」로 나타났다 —
  균등 배치 시절에는 라이저 길이가 전부 같은 값으로 나갔기 때문이다.

■ 규범

  좌표를 믿는 조건 = 거의 모든(95%) 배관에서 비가 [0.9, 1.1] 안.
  그때 덮어쓰기는 미세 보정이지 교체가 아니다. 흩어져 있으면 **선언이 권위**.
  (80% 로 느슨히 걸면 결합 SDF 가 여전히 뚫린다 — 맞는 다수 뒤에서 틀린
   11개가 27.6 m 를 지웠다. 그 실측도 여기 적는다.)
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _sdf(nodes, pipes) -> str:
    """최소 SDF — (label, x, y, elev) · (label, in, out, length)."""
    ns = "".join(
        f'<Node label="{lab}" elevation="{e}"><Position x="{x}" y="{y}"/>'
        f"</Node>" for lab, x, y, e in nodes)
    ps = "".join(
        f'<Pipe label="{lab}" input="{a}" output="{b}" length="{ln}"'
        f' bore="65" roughness-or-c="120" status="normal"/>'
        for lab, a, b, ln in pipes)
    # ★배관은 <Pipe-set> «안» 에 있어야 한다 — 파서는 Links 의 직접 자식
    #   <Pipe> 를 걷지 않는다(처음 그렇게 썼다가 셋 다 빈 망으로 떨어졌다).
    return (f'<?xml version="1.0"?><Network-spray>'
            f"<Nodes>{ns}</Nodes><Links><Pipe-set>{ps}</Pipe-set></Links>"
            f"</Network-spray>")


def _parse(nodes, pipes):
    from kfp_sdf_converter import parse_sdf
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "x.sdf"
        p.write_text(_sdf(nodes, pipes), encoding="utf-8")
        net = parse_sdf(str(p))
    lab_of = {}
    for pid, cp in net.pipes.items():
        lab_of[str((cp.raw or {}).get("sdf_label") or pid)] = float(
            cp.length_m or 0)
    return lab_of


def test_실미터_좌표는_믿는다():
    """KFP→SDF 왕복(좌표거리==length)은 비가 전부 1.000 — 좌표가 이긴다."""
    nodes = [("1", 0.0, 0.0, 0.0), ("2", 3.0, 0.0, 0.0),
             ("3", 3.0, 4.0, 0.0), ("4", 8.0, 4.0, 0.0)]
    pipes = [("P1", "1", "2", 3.001), ("P2", "2", "3", 3.999),
             ("P3", "3", "4", 5.002)]
    got = _parse(nodes, pipes)
    assert abs(got["P1"] - 3.0) < 1e-6      # 좌표거리로 미세 보정
    assert abs(got["P2"] - 4.0) < 1e-6
    assert abs(got["P3"] - 5.0) < 1e-6


def test_스키매틱_좌표는_선언을_덮지_못한다():
    """★실제 사고의 재현 — 다수가 우연히 1 근처, 소수가 크게 어긋난다.

    종전 중앙값 게이트는 이 망을 통과시켜 9.7 m 를 0.05 m 로 지웠다.
    """
    # 배관 20개: 16개는 좌표거리 == 선언(비 1.0), 4개는 크게 어긋남.
    nodes = [("n0", 0.0, 0.0, 0.0)]
    pipes = []
    x = 0.0
    for i in range(16):
        x2 = x + 2.0
        nodes.append((f"n{i + 1}", x2, 0.0, 0.0))
        pipes.append((f"ok{i}", f"n{i}", f"n{i + 1}", 2.0))
        x = x2
    # 어긋나는 4개 — 선언은 길지만 그림은 짧다(스키매틱 압축).
    prev = "n16"
    for i, ln in enumerate((9.735, 0.937, 10.346, 0.522)):
        lab = f"s{i}"
        nodes.append((lab, x + 0.05 * (i + 1), 0.0, 0.0))
        pipes.append((f"bad{i}", prev, lab, ln))
        prev = lab
    got = _parse(nodes, pipes)
    assert abs(got["bad0"] - 9.735) < 1e-6, "선언 9.735 m 가 좌표에 지워졌다"
    assert abs(got["bad1"] - 0.937) < 1e-6
    assert abs(got["bad2"] - 10.346) < 1e-6
    assert abs(got["bad3"] - 0.522) < 1e-6
    # 일치하던 다수도 선언 그대로다(2.0) — 반쪽만 덮는 일은 없다.
    assert abs(got["ok0"] - 2.0) < 1e-6


def test_총연장이_보존된다():
    """S443 — 세 형식 대조가 «참값» 으로 일치해야 뜻이 있다.

    종전에는 세 형식이 같은 오염원(좌표)에서 나와 서로 일치했고, 그래서
    cross_check 가 오염을 잡지 못했다. 선언 총합이 권위다.
    """
    nodes = [("n0", 0.0, 0.0, 0.0)]
    pipes = []
    x = 0.0
    for i in range(10):
        x2 = x + 1.0
        nodes.append((f"n{i + 1}", x2, 0.0, 0.0))
        pipes.append((f"p{i}", f"n{i}", f"n{i + 1}", 1.0 if i < 8 else 7.0))
        x = x2
    got = _parse(nodes, pipes)
    assert abs(sum(got.values()) - (8 * 1.0 + 2 * 7.0)) < 1e-6
