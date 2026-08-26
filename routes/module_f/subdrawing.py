# -*- coding: utf-8 -*-
"""[H-2 · H-3] 계통도 · 기계실 슬롯의 엔진 — 모듈 A 를 그대로 문다.

특허 S650 은 «평면도 · 계통도 · 기계실에 같은 절차를 반복 적용» 하라고 한다.
그런데 «같은 절차» 가 «같은 구현» 을 뜻하지는 않는다. 세 도면은 제1국면에서
읽는 것이 다르다.

  평면도   배관 + 헤드 기호 → 사람이 재료를 찍고(E) 손질한 board 위에서 최불리
  계통도   층을 세로로 늘어놓은 모식도 → 펌프·알람밸브 두 점 사이의 «경로»
  기계실   수원(탱크) → 입상관 연결점 사이의 «경로» (평면 좌표 보존)

뒤의 둘은 board 를 만들 이유가 없다 — 찍을 재료가 아니라 두 점을 잇는 경로가
전부다. 모듈 A 가 이 둘을 이미 정확히 그렇게 뽑는다(`extract_system_path` ·
`extract_machine_room_path` — A 의 주석이 «동형, 같은 그래프/Dijkstra» 라고
적어 뒀다). **새로 짜지 않고 그대로 부른다.**

여기서 하는 일은 둘이다:
  ① A 의 entity 목록을 F 캔버스가 읽는 World 모양으로 옮긴다(어댑터)
  ② A 의 두 추출기를 F 의 세션 규약으로 감싼다

★두 점은 사람이 찍는다. 특허 S220 의 우선순위(사용자 지정 → 자동 탐지 →
  최다 접속 절점)에서 F 는 첫째만 쓴다 — 계통도의 펌프·알람밸브는 도면마다
  기호가 달라 자동 탐지가 조용히 틀리면 경로가 통째로 다른 곳으로 간다.
"""
from __future__ import annotations

# A 의 entity 는 색을 싣지 않는다(레이어만). 캔버스는 레이어×색으로 묶으므로
# 색을 하나로 고정하면 «레이어 단위» 묶음이 된다 — 계통도에는 그게 맞다.
_COLOR = 7
# 텍스트 높이 — A 의 entity 에는 없다. 0 이면 `_world_payload` 가 아니라
# 치수 텍스트 판독(`extract_dia_text_points`)에서 걸리므로 양수를 준다.
_TEXT_H = 1.0


class _EntWorld:
    """`_world_payload` 가 읽는 World 의 최소 모양.

    stage1 의 World 와 필드 이름·튜플 모양이 같아야 한다 — 그래야 평면도와
    같은 캔버스 코드로 그려지고, 레이어 토글·묶음 통계가 공짜로 따라온다.
    """

    __slots__ = ("segs", "raw_segs", "circles", "arcs", "arc_ang", "texts")

    def __init__(self):
        self.segs = []       # (layer, color, (x1,y1), (x2,y2))
        self.raw_segs = []
        self.circles = []    # (layer, color, cx, cy, r)
        self.arcs = []       # (layer, color, cx, cy, r)
        self.arc_ang = []    # arcs 와 같은 순 · (start, sweep)
        self.texts = []      # (layer, color, x, y, h, s)


def entities_to_world(entities) -> _EntWorld:
    """A 의 entity 목록 → 캔버스가 그릴 수 있는 World.

    폴리선은 마디마다 선분으로 편다 — 캔버스가 선분만 그리기 때문이고,
    계통도의 배관은 어차피 마디 단위로 잰다.
    """
    w = _EntWorld()
    for en in (entities or ()):
        t = en.get("t")
        lay = str(en.get("l") or "0")
        if t == "L":
            p = en.get("p") or []
            if len(p) >= 4:
                w.segs.append((lay, _COLOR, (float(p[0]), float(p[1])),
                               (float(p[2]), float(p[3]))))
        elif t == "PL":
            pts = en.get("p") or []
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                if len(a) >= 2 and len(b) >= 2:
                    w.segs.append((lay, _COLOR, (float(a[0]), float(a[1])),
                                   (float(b[0]), float(b[1]))))
        elif t == "C":
            c = en.get("c") or []
            if len(c) >= 2:
                w.circles.append((lay, _COLOR, float(c[0]), float(c[1]),
                                  float(en.get("r") or 0.0)))
        elif t == "A":
            c = en.get("c") or []
            if len(c) >= 2:
                w.arcs.append((lay, _COLOR, float(c[0]), float(c[1]),
                               float(en.get("r") or 0.0)))
                a = en.get("a") or [0.0, 360.0]
                sa = float(a[0])
                ea = float(a[1]) if len(a) > 1 else sa + 360.0
                sweep = (ea - sa) % 360.0 or 360.0
                w.arc_ang.append((sa, sweep))
        elif t == "T":
            p = en.get("p") or []
            if len(p) >= 2:
                w.texts.append((lay, _COLOR, float(p[0]), float(p[1]),
                                _TEXT_H, str(en.get("v") or "")))
    return w


def parse_subdrawing(dxf_path):
    """계통도·기계실 DXF 한 장을 읽는다 — A 의 시각화 우선 파서 그대로.

    `include_hidden_layers=True` 다. 계통도는 꺼둔 레이어에 배관이 있는 일이
    흔해서(A 의 주석), 숨긴 것을 빼면 경로가 끊긴다.
    """
    from remote30_prototype import parse_dxf_for_view
    parsed = parse_dxf_for_view(dxf_path, include_hidden_layers=True)
    return parsed.get("entities") or [], parsed


def extract_system(entities, pump_xy, av_xy, *, snap_tolerance_mm=2500.0,
                   waypoints=None, floor_profile_rows=None) -> dict:
    """S720 — 계통도에서 펌프 → 알람밸브 경로(입상관)를 뽑는다.

    실패(클릭이 배관에서 너무 멀다 · 두 점이 안 이어진다)는 `ValueError` 로
    올라온다. 임의로 메우지 않는다 — 특허 S340 의 «실제 배관만으로 이어지지
    아니하는 자리는 임의로 메우지 아니하고 미도달로 보고한다» 를 따른다.
    """
    from remote30_prototype import extract_system_path
    return extract_system_path(
        entities, tuple(pump_xy), tuple(av_xy),
        snap_tolerance_mm=float(snap_tolerance_mm),
        waypoints=waypoints or None,
        floor_profile_rows=floor_profile_rows or None)


def extract_system_clean(dxf_path, *, scale_mm_per_unit: float = 1.0) -> dict:
    """조각난 풀 계통도 대신 «깨끗한 배관망 DXF» 를 통째로 읽는 폴백.

    실측(계통도_LH_306): 풀 도면의 PIPE 레이어는 단일망 추출이 안 된다 —
    강제 봉합으로 경로가 엉뚱한 데로 튄다. 손으로 정리한 배관망 파일은 단일
    연결망이라 추측 bridge 없이 그대로 읽힌다.
    """
    from remote30_prototype import extract_clean_system_network
    return extract_clean_system_network(dxf_path,
                                        scale_mm_per_unit=float(scale_mm_per_unit))


def extract_machineroom(entities, source_xy, conn_xy, *,
                        snap_tolerance_mm=2500.0, ceiling_m=None) -> dict:
    """S730 — 기계실에서 수원(탱크) → 입상관 연결점 경로를 뽑는다.

    좌표는 **평면 그대로 보존**된다(H-D6). 계통도처럼 세로 막대로 재배치하면
    기계실 배관의 평면 형상이 뭉개진다 — A 도 그래서 둘을 갈라 놓았다.

    라벨은 `m1..mK` 라 라이저(1~10)·헤드(10+)와 부딪히지 않는다.
    """
    from remote30_prototype import extract_machine_room_path
    return extract_machine_room_path(
        entities, tuple(source_xy), tuple(conn_xy),
        snap_tolerance_mm=float(snap_tolerance_mm),
        ceiling_m=(float(ceiling_m) if ceiling_m is not None else None))


def riser_summary(riser: dict) -> dict:
    """추출 결과 한 줄 — 화면과 슬롯 탭이 같은 말을 하게 한다."""
    r = riser or {}
    nodes = r.get("nodes") or []
    pipes = r.get("pipes") or []
    total_m = 0.0
    for p in pipes:
        try:
            total_m += float(p.get("length") or p.get("length_m") or 0.0)
        except (TypeError, ValueError):
            pass
    return {
        "nodes": len(nodes),
        "pipes": len(pipes),
        "total_m": round(total_m, 2),
        "av_node_label": r.get("av_node_label"),
        "source_node_label": r.get("source_node_label"),
        "conn_node_label": r.get("conn_node_label"),
        # 추측으로 이은 자리 — 화면이 점선으로 갈라 그려야 한다.
        "bridges": r.get("bridge_count") or r.get("bridges") or 0,
    }
