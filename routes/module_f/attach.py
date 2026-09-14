# -*- coding: utf-8 -*-
"""[전개가 붙일 수 있는 헤드] 한 판에 한 번만 재고 두 화면이 나눠 쓴다.

■ 왜 필요한가 (2026-09-09 사용자 지적)

  「애초에 고른다는 선택지 없이, 사전에 체크한 30개의 헤드와 거기 연결된
   배관만 가져오면 되는 것 아닌가」

  옳은 지적이고, 코드는 절반쯤 그렇게 되어 있었는데 **순서가 거꾸로**였다::

      손질     최불리 K개를 고른다            ← board 도달로만 센다
      수리계산 cand = 고른 것 ∩ 붙는 헤드     ← 여기서 처음 걸러진다
               worst_k_heads(k=K, only=cand)  ← 그래서 K 보다 적게 나온다

  실측(대명동 골든 K=10): 손질이 10개를 골랐는데 그중 2개를 전개가 배관에
  붙이지 못해 표에는 **8개**만 왔다. 손질이 고를 때 이미 「붙는 헤드」만
  후보로 삼았다면 10개를 골랐을 것이고 그대로 표에 왔을 것이다.

■ 왜 «한 번만» 인가

  `attachable_heads` 는 **전체망 전개를 통째로 한 번** 돈다. 종전에는
  수리계산을 누를 때마다 돌았다(캐시 없음) — 사용자가 지적한 「지하주차장
  도면에서 수리계산이 너무 오래 걸린다」의 큰 몫이 여기다.

  판(board)이 안 바뀌었으면 답도 안 바뀐다. 그래서 판의 «지문» 으로 캐시한다.
  손질이 한 번 내고 수리계산이 그대로 받아 쓰므로, 두 화면을 합쳐도 **판마다
  한 번**이다 — 종전(수리계산마다 한 번)보다 적게 든다.

★지문은 «판이 달라지면 반드시 달라져야» 한다. 하나라도 빠뜨리면 옛 답을
  새 판에 쓰게 되고, 그것은 조용한 오답이다 — 이 저장소가 옛 최불리 선정에서
  이미 겪은 함정이다(`4a13d21`).
"""
from __future__ import annotations


def board_stamp(es) -> tuple:
    """판의 지문 — 이것이 같으면 «붙는 헤드» 도 같다.

    ★값이 아니라 **판을 바꾸는 모든 손잡이**를 넣는다. 절점·간선·헤드는
      개수와 함께 마지막 좌표까지 본다(개수가 같은데 자리만 옮기는 편집이
      있다). 급수원·밸브는 자리 자체가 답을 바꾸므로 통째로 넣는다.
    """
    b = es.board
    pts = getattr(b, "pts", None) or ()
    edges = getattr(b, "edges", None) or ()
    disks = getattr(b, "disks", None) or ()
    return (
        getattr(es, "key", None),
        len(pts), len(edges), len(disks),
        tuple(getattr(b, "sources", None) or ()),
        tuple(getattr(b, "valves", None) or ()),
        tuple(getattr(b, "disk_kinds", None) or ()),
        len(getattr(b, "joins", None) or ()),
        len(getattr(b, "deletes", None) or ()),
        len(getattr(b, "edge_len_mm", None) or {}),
        # 개수가 같아도 자리가 바뀌었으면 다른 판이다.
        tuple(tuple(round(float(v), 1) for v in p) for p in pts[-3:]),
        tuple(tuple(round(float(v), 1) for v in d) for d in disks[-3:]),
    )


def wet_heads(sess, es, *, selected_source=None):
    """전개가 배관에 붙일 수 있는 헤드 — 판마다 한 번만 잰다.

    반환은 `design.restrict.attachable_heads` 의 것 그대로:
    `{"ok", "wet": set(hcov 번호), "total", "dropped"}`.
    실패해도 그대로 돌려준다 — 부르는 쪽이 «못 쟀다» 를 알아야 한다.
    """
    from services.cad_import.design.restrict import attachable_heads

    stamp = (board_stamp(es), selected_source)
    hit = sess.get("_wet_probe")
    if hit is not None and hit.get("stamp") == stamp:
        print("[붙는 헤드] 판이 그대로라 다시 재지 않습니다"
              f" — 붙는 헤드 {len(hit['probe'].get('wet') or ())}개")
        return hit["probe"]
    # ★왜 다시 재는지 말한다. 「아끼려고 캐시를 뒀는데 매번 빗나간다」가
    #   조용히 벌어지면 비용만 늘고 아무도 모른다 — 실제로 한 번 그랬다.
    why = ("처음" if hit is None else
           ("급수원이 바뀜" if hit["stamp"][0] == stamp[0] else "판이 바뀜"))
    print(f"[붙는 헤드] 전체망 전개를 한 번 돕니다 ({why})…")
    probe = attachable_heads(es.convert_payload(),
                             selected_source=selected_source)
    sess["_wet_probe"] = {"stamp": stamp, "probe": probe}
    return probe

def picked_heads_wet(es, picked, *, selected_source=None):
    """**뽑힌 K 개만** 전개해 「배관에 붙는가」를 잰다 — 전체망을 안 돈다.

    ■ 왜 (2026-09-14 사용자 지적: B1F 최불리가 너무 오래 걸린다)

      시계를 대 보니 답이 한 줄이었다 — B1F K=30 실측::

          판 열기                    0.80s
          ★전체망 전개(wet_heads)  154.10s   98.2%
          최불리 선정(Dijkstra)      0.44s
          먼 순서 검사(K+1)          0.51s
          ★제한 전개(K개만)          1.13s   ← 같은 답을 여기서 얻는다

      `/edit/worst` 가 «붙는 헤드» 를 재는 이유는 §2-3 하나다 — **뽑힌 K 개**
      중 못 붙는 것이 있으면 막는다. 그 판정에 도면 전체 3,235개를 전개할
      이유가 없다. 30개만 전개하면 **1.13초**고 답은 같다.

      잃는 것: 「이 도면에 안 붙는 헤드가 N개 있다」는 전체 통계. 그것은
      수리계산(`/design/build`)이 «제외 사유» 로 여전히 낸다 — 그쪽은 진행
      표시가 있는 잡이라 오래 걸려도 화면이 얼지 않는다.

    반환은 `wet_heads` 와 같은 모양이되 번호는 **board 헤드 번호**다
    (`restrict_to_worst` 가 만든 지역 번호를 되돌려 준다).
    """
    from services.cad_import.convert.planar import build_planar_graph
    from services.cad_import.design.restrict import restrict_to_worst

    order = sorted({int(i) for i in (picked or ())})
    if not order:
        return {"ok": False, "error": "뽑힌 헤드가 없습니다.",
                "wet": set(), "total": 0, "dropped": 0,
                "reason": {}, "shared": set()}
    payload = es.convert_payload()
    lim = restrict_to_worst(payload, es.board, {"heads": order})
    built = build_planar_graph(
        lim.get("key") or "picked", write=False,
        selected_source=selected_source or lim.get("selected_source"),
        pts=lim.get("pts"), edges=lim.get("edges"), hcov=lim.get("hcov"),
        ups=lim.get("ups"), head_kinds=lim.get("head_kinds"),
        user_sources=lim.get("sources"), ho=lim.get("ho"),
        edge_len_mm=lim.get("edge_len_mm"))
    if not built.get("ok"):
        return {"ok": False,
                "error": built.get("error") or "제한 전개가 실패했습니다.",
                "wet": set(), "total": len(order), "dropped": 0,
                "reason": {}, "shared": set()}

    def back(i):
        i = int(i)
        return order[i] if 0 <= i < len(order) else None

    wet = {b for b in (back(i) for i in (built.get("wet_head_idx") or ()))
           if b is not None}
    reason = {}
    for i, why in (built.get("head_reason") or {}).items():
        b = back(i)
        if b is not None:
            reason[b] = why
    shared = {b for b in (back(i) for i in (built.get("shared_head_idx") or ()))
              if b is not None}
    print(f"[붙는 헤드] 뽑힌 {len(order)}개만 전개했습니다 — 붙는 것 {len(wet)}개"
          f" (전체망을 돌지 않습니다)")
    return {"ok": True, "wet": wet, "total": len(order),
            "dropped": len(order) - len(wet),
            "reason": reason, "shared": shared}
