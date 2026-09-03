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
