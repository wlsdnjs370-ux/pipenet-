# -*- coding: utf-8 -*-
"""접속점(앵커) — 헤드망이 라이저와 만나는 그 한 점. 물이 들어오는 자리다.

■ 낱말을 여기서 못박는다

  이 저장소에서 «앵커» 는 오랫동안 **정반대 두 곳**을 가리켰다.
    ① 라이저와 헤드망이 만나는 접합점(= 알람밸브 자리)
       — `merge.ANCHOR_LABEL="10"` · `/api/module-f/auto/anchor`(알람밸브 위치)
       · `/api/module-f/edit/anchor-click`(알람밸브 원클릭)
    ② 급수원에서 가장 먼 헤드(기준압을 잡는 지점) — 옛 `worst["anchor"]`
  망의 **입구**와 가장 먼 **출구**를 같은 이름으로 부르면, 읽는 사람이 매번
  어느 쪽인지 되짚어야 한다. 수리계산에서 그 둘은 가장 헷갈리면 안 되는 두
  지점이다.

  → 「앵커」는 이제 ① 하나만 뜻한다. ② 는 «기준 헤드»(`worst["worst_head"]`)다.

■ 왜 «펌프» 라는 이름이 남아 있나

  접속점 노드는 .kfp 에 `type_id="pump"` 로 저장된다. 실체는 펌프가 아니라
  **라이저 접속점(알람밸브 자리)** 이다 — 진짜 수원·펌프는 기계실에서 오고,
  통합(S740)하는 순간 이 노드는 라이저의 AV 노드에 자리를 내주고 사라진다
  (`remote30_full_network`: 「AV 는 라이저 쪽에서 이미 포함 — 헤드망 쪽 사본
  skip」). 저장된 이름을 바꾸면 이미 만들어 둔 .kfp 가 안 열리므로, 데이터는
  그대로 두고 **부르는 이름만** 바로잡는다.

■ 못 찾으면 조용히 넘어가지 않는다

  종전에는 두 곳(`design/tables.py` · `design/restrict.py`)이 똑같이
  「접속점을 못 찾으면 `next(iter(nodes))`」 로 눕었다 — **dict 에 먼저 들어온
  노드가 Input 경계**가 되고, 거기서 물 흐르는 방향이 통째로 다시 유도된다.
  사람이 찍은 자리와 아무 상관이 없는 지점에서 계산이 시작되는 것이라, 나온
  표는 «틀렸다» 고 말해 주지도 않으면서 틀린다. 그래서 여기서는 던진다.
"""
from __future__ import annotations

# 저장된 표식 — 역사적 이름이다(위 「왜 «펌프»」 참고). 값을 바꾸면 기존
# .kfp 가 접속점을 잃으므로 바꾸지 않는다.
ANCHOR_TYPE_ID = "pump"


class AnchorMissing(ValueError):
    """접속점이 없다 — 어디서 물이 들어오는지 모르는 망이다."""


def find_anchor(meta_nodes) -> str | None:
    """접속점 노드 id. 없으면 None — 부르는 쪽이 «없음» 을 반드시 다뤄야 한다."""
    for nid, meta in (meta_nodes or {}).items():
        if str((meta or {}).get("type_id", "")) == ANCHOR_TYPE_ID:
            return nid
    return None


def require_anchor(meta_nodes, *, what: str = "이 망") -> str:
    """접속점 노드 id. 없으면 사람이 읽을 문장으로 던진다.

    ★아무 노드나 골라 잇는 «폴백» 을 두지 않는다. 그 폴백은 실패를 감추기만
      하고, 감춘 채로 나온 표는 Input 위치도 물 흐르는 방향도 틀린다.
    """
    nid = find_anchor(meta_nodes)
    if nid is not None:
        return nid
    n = len(meta_nodes or {})
    raise AnchorMissing(
        f"{what}에 접속점(알람밸브 자리)이 없습니다 — 물이 어디로 들어오는지 "
        f"모르는 상태라 표를 만들 수 없습니다. 손질 단계에서 알람밸브를 "
        f"찍으세요. (노드 {n}개 중 type_id='{ANCHOR_TYPE_ID}' 없음)")


# 부착점 되짚기(tables.WORST_HEAD_SNAP_M)와 같은 근거의 허용치 — 실측(B1F)으로
# 노드정리의 «직선 위치 복원» 이 좌표를 6~8mm 옮긴다. 100mm 는 그 열 배가 넘고
# 이웃 배관 노드 간격에는 한참 못 미친다.
VALVE_SNAP_M = 0.10


def valve_kfp_nodes(meta_nodes, board_pts, valve_board_nodes,
                    origin_mm) -> list:
    """board 알람밸브 픽 → kfp 노드 id 들.

    ★§29 실측 — 이 대응이 없어서 사람이 찍은 알람밸브가 산출물(기기표)에
      한 번도 실리지 못했다. `build_design_tables(valve_nodes=…)` 자리는 처음부터
      있었는데 **아무도 안 넘겼다**(웹도, 데스크톱 G 창의 명시적 None 도).

    되짚기는 좌표다 — `node_ref` 류의 역참조 표는 노드정리 «전» 을 가리켜 절반
    넘게 틀린다(§30 실측 30개 중 12개). 리팩터링 7 이후 알람밸브 픽은 접속점을
    겸하므로 보통 pump 도장 노드에 정확히 떨어지지만, 옛 저장본(밸브만 따로
    찍힌 것)도 같은 식으로 이어진다.

    못 이은 픽은 조용히 버리지 않고 목록 밖으로 남긴다 — 부르는 쪽이 개수를
    맞대 보고 사람에게 말할 수 있게, (이은 것, 못 이은 board 노드) 를 준다.
    """
    import math

    hit: list = []
    missed: list = []
    if origin_mm is None or not board_pts:
        return [], [n for n in (valve_board_nodes or ())]
    for bn in (valve_board_nodes or ()):
        if not isinstance(bn, int) or not (0 <= bn < len(board_pts)):
            missed.append(bn)
            continue
        tx = (float(board_pts[bn][0]) - float(origin_mm[0])) / 1000.0 + 1.0
        ty = (float(board_pts[bn][1]) - float(origin_mm[1])) / 1000.0 + 1.0
        best = None
        for nid, meta in (meta_nodes or {}).items():
            tid = str((meta or {}).get("type_id", ""))
            if tid == "head":
                continue          # 알람밸브가 헤드일 수는 없다
            c = (meta or {}).get("coords") or (0.0, 0.0, 0.0)
            d = math.hypot(float(c[0]) - tx, float(c[1]) - ty)
            if d > VALVE_SNAP_M:
                continue
            # ★허용 안에 접속점이 있으면 **그것이** 그 밸브다. 알람밸브 픽은
            #   접속점을 겸하기 때문이다(리팩터링 7). 거리만으로 고르면 안 된다:
            #   전개가 그 자리에 곁가지 노드를 만들어 두면 몇 mm 차이로 그쪽이
            #   이기는데, 그 곁가지는 담당 헤드 0 인 막다른 관이라 기기가
            #   «물이 안 지나는 관» 에 붙는다(실측 B1F: 접속점에서 17mm 떨어진
            #   base 노드가 이겨 25A 곁가지에 붙었고, 라이브러리에 25A 알람밸브가
            #   없어 등가길이가 미해결로 떨어졌다).
            rank = (0 if tid == ANCHOR_TYPE_ID else 1, d)
            if best is None or rank < best[0]:
                best = (rank, nid)
        if best is None:
            missed.append(bn)
        elif best[1] not in hit:
            hit.append(best[1])
    return hit, missed
