# -*- coding: utf-8 -*-
"""헤드 연결 검사 — 선택한 K개는 표 확정, 전체망은 별도 진단.

손질의 최불리 선정과 수리계산 표 확정은 선택한 헤드만 전개하여 연결을
확인한다. 연결되지 않은 선정 헤드를 다른 후보로 교체하지 않는다.
전체 도면의 제외 사유는 /design/diagnose 요청 때 검사하고 판 지문으로
캐시한다. 선정 없이 직접 호출하는 옛 경로만 전체망 검사를 유지한다.
"""
from __future__ import annotations


def board_stamp(es) -> tuple:
    """판의 지문 — 이것이 같으면 «붙는 헤드» 도 같다.

    ★값이 아니라 **판을 바꾸는 모든 손잡이**를 넣는다. 절점·간선·헤드는
      개수와 함께 모든 좌표·간선을 본다(개수가 같은데 자리만 옮기는 편집이
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
        # Counts and the last three points miss edits in the middle of a
        # drawing. Diagnostic cache identity must cover the entire geometry.
        tuple(tuple(p) for p in pts),
        tuple(sorted(tuple(e) for e in edges)),
        tuple(tuple(d) for d in disks),
        tuple(sorted((getattr(b, "edge_len_mm", None) or {}).items())),
        repr(getattr(b, "head_kinds", None)),
        repr(getattr(b, "ups", None)),
        repr(getattr(b, "ho", None)),
    )


def wet_heads(sess, es, *, selected_source=None):
    """전개가 배관에 붙일 수 있는 헤드(도면 전체) — 판마다 한 번만 잰다.

    별도 전체 진단 또는 선정 없는 옛 호출에서 사용한다. 표 확정의 정상
    경로는 design_probe 로 선택된 헤드만 검사한다.

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

def design_probe(sess: dict, es, picked, *, selected_source=None) -> dict:
    """Check selected heads without expanding the rest of the drawing.

    Legacy callers without a selection retain the whole-network path. A
    selected probe cannot classify unselected heads as detached.
    """
    if not picked:
        return wet_heads(sess, es, selected_source=selected_source)
    probe = picked_heads_wet(es, picked, selected_source=selected_source)
    return dict(probe, scope="selected")


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
      별도 전체 도면 진단(`/design/diagnose`)에서 여전히 확인할 수 있다.

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
