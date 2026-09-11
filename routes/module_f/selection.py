# -*- coding: utf-8 -*-
"""선정 인계 · 표 신선도 — 라우트가 아닌 «판단» 만 사는 곳.

2026-09-11 리팩터링으로 api_design.py 에서 잘라 옮겼다(재타이핑 0 —
`data/_refactor_selection_split.py` 가 절단기). 두 가지 이유다:

  · api_design.py 가 1,597줄까지 자랐다 — 이 주에 자란 것이 전부 이 덩이다
    (두 화면 선정일치 §2-1~§2-5 · 영역 가두기 · 표 신선도).
  · `remote30._design_corridor` 가 이 판단들을 쓰려고 **라우트 파일을**
    import 하고 있었다. 순수 판단이 라우트 등록 파일에 살면 그런 의존이 는다.

여기 있는 것은 전부 순수 함수다 — Flask 도, 엔진 부팅(_boot)도 모듈 레벨에서
끌지 않는다(엔진 참조는 함수 안 지연 import 뿐). 그래서 시험이 서버 없이 바로
import 해 쓴다.

★이름 계약: api_design.py 가 이 이름들을 **다시 내보낸다** — 시험·탐침이
  `routes.module_f.api_design` 에서 import 하기 때문이다. 가른 것 때문에
  도구가 깨지면 안 된다(패키지 분리 때의 `__init__.py` 재수출과 같은 규약).
"""
from __future__ import annotations


def _load_map(got: dict) -> dict:
    """배관 id → **담당 헤드 수**. 간선 굵기의 근거다.

    두 곳이 안다: 도면 간선에서 온 배관은 `worst.loads`(board 간선 키), 세로
    구간처럼 역참조가 없는 배관은 `tree_loads`(배관 id 키).

    ★못 찾았을 때 **0 을 적지 않는다.** 종전에는 `loads.get(key, 0)` 으로 0 을
      박아 넣었고, 뒤따르는 `tree_loads` 폴백이 `setdefault` 라 영영 못 들어왔다.
      그래서 담당 헤드 수를 잃은 배관이 굵기 0 으로 그려졌다 — 주배관이 가지
      끝처럼 가늘게 보여, 아이소가 평면과 «위상이 달라 보이는» 원인이 됐다.

      실측(대명동 K30 · `scripts/_probe_design_topology.py`):
        edge_ref 216개 중 49개(22%)가 빗나갔고 tree_loads 가 그 49개를 전부
        알고 있었다. 고친 뒤 굵기 0 인 배관은 51 → 2 로 떨어졌다.
        빗나가는 이유는 «다른 길» 이 아니다 — 그 간선들은 corridor 선 위에
        거리 0mm 로 얹혀 있다. 평면 그래프를 세우며 티 겹침 정규화·직선 위치
        복원·노드정리가 절점을 다시 매겨 (i,j) 키만 어긋난 것이다.
    """
    loads = ((got.get("worst") or {}).get("loads")) or {}
    ref = got.get("edge_ref") or {}
    tree = got.get("tree_loads") or {}
    out: dict = {}
    for pid, edge in ref.items():
        try:
            i, j = int(edge[0]), int(edge[1])
        except (TypeError, ValueError, IndexError):
            continue
        v = loads.get((min(i, j), max(i, j)))
        if v is not None:
            out[str(pid)] = int(v)
    for pid, n in tree.items():
        try:
            out.setdefault(str(pid), int(n))
        except (TypeError, ValueError):
            continue
    return out


def _head_row(i, board, reasons=None) -> dict:
    """헤드 한 개를 «사람이 찾아갈 수 있는 한 줄» 로 — 번호·자리·사유.

    번호만 주면 아무도 그 헤드를 도면에서 못 찾는다. 사유는 전개가 이미 가른
    것을 옮겨 적을 뿐이다(§2-4) — 여기서 판정하지 않는다.
    """
    row: dict = {"disk": int(i)}
    disks = list(getattr(board, "disks", None) or ())
    if 0 <= int(i) < len(disks):
        d = disks[int(i)]
        row["xy"] = [round(float(d[0]), 1), round(float(d[1]), 1)]
    why = (reasons or {}).get(int(i))
    if why is None:
        why = (reasons or {}).get(str(i))
    if why:
        row["why"] = str(why)
        try:
            from services.cad_import.convert.planar import HEAD_REASON_TEXT
            row["why_text"] = HEAD_REASON_TEXT.get(str(why), str(why))
        except Exception:  # noqa: BLE001 — 문구가 없어도 갈래는 말한다
            row["why_text"] = str(why)
    return row


def _in_rect(p, r) -> bool:
    return (float(r[0]) <= float(p[0]) <= float(r[2])
            and float(r[1]) <= float(p[1]) <= float(r[3]))


def zone_confined_pool(pool, picked, zones, disks):
    """[영역 넘어감] 채울 후보를 «선정이 이미 든 영역» 안으로 가둔다.

    ■ 증상 (2026-09-11 사용자 지적 · 대명동)

      「영역1 내에 있는 2개 헤드가 영역2 및 배관망을 강제로 넘어간다.」

      실측으로 재현했다(`data/_probe_zone_cross.py` · 영역1 을 K=10 에 딱 맞게
      좁힌 판)::

          손질 선정 10개        전부 영역1
          못 붙는 헤드 2개      123·122 (pass_under)
          채움                  헤드 47 → 영역1 · 헤드 82 → **영역2**
          표 최종               영역1 9 · **영역2 1**
          corridor              두 영역에 걸침 · 배관 78 → **120**

    ■ 왜 넘어갔나

      채움 후보(`worst_cand`)가 **영역 합집합**이다. 영역1 에서 못 붙은
      자리를 «그다음으로 불리한 헤드» 로 채우는데, 그 헤드가 영역2 에 있다.
      먼 순서 하나로만 고르니 영역 경계를 모른다.

    ■ 왜 그것이 틀렸나

      설계면적은 «하나의 방호구역 안에서 **인접한** K개» 다(NFPC). 두 영역에
      걸친 K개는 동시에 방수되는 무리가 아니고, corridor 도 건물을 가로질러
      관경·유하거리가 실제와 달라진다. 사람이 사각형을 **둘로 나눠** 그린 것은
      「이 둘은 다른 구역」이라는 뜻이다 — 합쳐 달라는 뜻이 아니다.

    ■ 규칙

      선정이 실제로 든 영역만 남긴다. 영역1 에서 뽑았으면 영역1 안에서만
      채운다. 두 영역에 걸쳐 뽑혔으면(사람이 그렇게 고른 것이다) 둘 다 남긴다.
      영역을 안 그렸으면(`zones` 없음) 도면 전체가 후보다 — **종전 그대로**다.

    반환: 가둔 후보 목록. 가둘 근거가 없으면 받은 것을 그대로 돌려준다.
    """
    rects = [r for r in (zones or ()) if r and len(r) >= 4]
    if not rects or not pool:
        return pool
    disks = list(disks or ())

    def at(i):
        return disks[int(i)] if 0 <= int(i) < len(disks) else None

    used = [ri for ri, r in enumerate(rects)
            if any(at(i) is not None and _in_rect(at(i), r) for i in picked)]
    if not used:
        # 선정이 어느 사각형에도 안 들어간다 — 가둘 근거가 없으니 손대지 않는다.
        return pool
    keep = [i for i in pool
            if at(i) is not None and any(_in_rect(at(i), rects[ri])
                                         for ri in used)]
    if len(keep) != len(pool):
        print(f"[영역] 채움 후보를 선정이 든 영역 {len(used)}곳 안으로 가둡니다"
              f" — {len(pool)} → {len(keep)}개 (설계면적은 한 구역 안이다).")
    return keep


# ★[두 화면 선정일치 §2-1] 손질이 덧붙였고 선정 계산은 모르는 칸.
#
#   `worst_k_heads` 가 내는 것은 «어느 헤드·어느 경로» 까지다. 어느 영역에서
#   어느 급수원 기준으로 뽑았는지는 손질 화면이 알고 덧붙인 값이라, 최종
#   선정으로 갈아 끼울 때 **함께 물려주지 않으면 화면이 조용히 다르게 그린다**
#   (영역 사각형이 사라지고 급수원 이름이 빈다).
_EDIT_ONLY_WORST_KEYS = ("zones", "sheet", "source_tag", "source_index",
                         "candidates")


def _adopt_final_worst(sess: dict, got: dict) -> dict | None:
    """[§2-1] 표에 **실제로 들어간 선정**을 세션의 선정으로 삼는다.

    두 화면이 다른 헤드를 그리던 원인은 선정이 두 벌이었기 때문이다:

        평면에서 보기   sess["worst"]        ← 손질이 고른 K개
        수리계산 표     got["worst"]         ← 못 붙는 것을 다음 순위로 채운 K개

    채우는 동작 자체는 옳다(기준개수 K를 지킨다 · `a44ec64`). 잘못은 그
    결과를 **아무도 화면에 돌려주지 않은 것**이다. 그래서 여기서 한 곳으로
    모은다 — 세울 계약은 한 줄이다:

        「평면에서 보기」가 그리는 corridor·선정 헤드는 표에 실제로 들어간
        선정과 언제나 같은 집합이다. 표가 아직 없으면 손질 선정을 그린다.

    ★손질 원본은 `sess["worst_edit"]` 에 **한 번만** 접어 둔다. 두 번째
      build 는 이미 최종 선정을 받으므로(멱등) 원본을 덮어쓰면 안 된다 —
      덮으면 「사람이 고른 것」이 영영 사라진다.
    ★`worst_rev` 를 지운다. `_edit_state` 는 지문이 같으면 corridor 를 안
      싣는데(1KB 규약), 안 지우면 화면이 옛 망을 그대로 들고 있는다.
    """
    fin = got.get("worst")
    if not isinstance(fin, dict) or not fin.get("heads"):
        return None
    prev = dict(sess.get("worst") or {})
    # 접어 두는 것은 «사람이 고른 것» 뿐이다. 두 번째 확정의 `prev` 는 이미
    # 표에서 온 선정이므로(from_design) 그것을 원본이라 부르면 안 된다.
    if prev and not prev.get("from_design") and not sess.get("worst_edit"):
        sess["worst_edit"] = prev
    base = sess.get("worst_edit") or prev
    out = dict(fin)
    for k in _EDIT_ONLY_WORST_KEYS:
        if k in base:
            out[k] = base[k]
    # ★«표에서 온 선정» 이라는 표시는 **선정 자신이** 든다. `worst_edit` 의
    #   유무로 미루면, 손질에서 최불리를 안 누른 세션(원본이 없다)에서 화면이
    #   「표와 같은 선정」을 그리면서 아니라고 말하게 된다.
    out["from_design"] = True
    sess["worst"] = out
    sess.pop("worst_rev", None)
    n_sw = len(set(prev.get("heads") or ()) - set(out.get("heads") or ()))
    if n_sw:
        print(f"[두 화면] 평면 보기의 선정을 표와 맞췄습니다 — 바뀐 헤드 {n_sw}개"
              f" (손질 원본은 그대로 두고 있습니다).")
    return out


def _selection_sig(sess: dict) -> tuple:
    """[표가 옛 것인가] «수리계산이 재료로 삼는 것» 의 지문.

    표(그리고 아이소)는 «표 확정» 을 누른 그 순간의 선정·손질판으로 만들어진
    사진이다. 그 뒤에 사람이 최불리를 다시 고르거나 손질을 고치면, 화면은
    평면 보기만 새 것으로 바뀌고 **아이소는 옛 표를 계속 그린다.**

    실측 2026-09-11: 최불리를 왼쪽 영역(K=8)으로 확정한 뒤 오른쪽 영역으로
    다시 고르면, 평면 corridor 는 새 8개인데 표 노즐은 옛 8개 그대로였다 —
    **겹치는 헤드 0개**. 화면은 아무 말도 안 했다. 사용자가 「최불리에서
    등록한 배관망이 아이소에 반영이 안 된다」고 한 그 자리다.

    지문에는 «표를 바꾸는 손잡이» 를 다 넣는다 — 하나라도 빠뜨리면 바뀐 판을
    안 바뀌었다고 말하게 되고, 그것은 조용한 오답이다.
    """
    w = sess.get("worst") or {}
    b = getattr(sess.get("edit"), "board", None)
    board_rev = (
        len(getattr(b, "pts", None) or ()),
        len(getattr(b, "edges", None) or ()),
        len(getattr(b, "disks", None) or ()),
        len(getattr(b, "joins", None) or ()),
        len(getattr(b, "deletes", None) or ()),
        tuple(getattr(b, "sources", None) or ()),
        tuple(getattr(b, "valves", None) or ()),
        tuple(getattr(b, "disk_kinds", None) or ()),
    )
    return (
        tuple(sorted(int(i) for i in (w.get("heads") or ()))),
        str(w.get("source_tag") or ""),
        w.get("sheet"),
        tuple(tuple(round(float(v), 1) for v in r) for r in (w.get("zones") or ())),
        board_rev,
    )


def _design_stale(sess: dict) -> dict | None:
    """지금 화면의 표·아이소가 옛 것이면 «무엇이 달라졌는지» 를 돌려준다.

    None 이면 최신이다. 조용히 두지 않는다 — 사람이 보는 그림이 제 결정과
    다른데 화면이 말하지 않으면, 그 그림을 믿고 다음 결정을 한다.
    """
    d = sess.get("design")
    if not d or d.get("sig") is None:
        return None
    old, now = d["sig"], _selection_sig(sess)
    if old == now:
        return None
    why = []
    if old[0] != now[0]:
        gone = len(set(old[0]) - set(now[0]))
        why.append(f"최불리 선정이 바뀌었습니다 (헤드 {len(old[0])} → "
                   f"{len(now[0])}개 · 그중 {gone}개가 다른 헤드로)")
    if old[1:4] != now[1:4]:
        why.append("영역·급수원·도면 장 중 하나가 바뀌었습니다")
    if old[4] != now[4]:
        why.append("손질에서 배관망을 고쳤습니다")
    return {"why": why or ["설정이 바뀌었습니다"],
            "heads_now": len(now[0]), "heads_in_table": len(old[0])}


def _worst_handoff_note(got: dict, picked: list, k_use: int, k_cfg: int,
                        filled: int = 0, board=None,
                        reasons: dict | None = None) -> dict:
    """[최불리 인계] 손질 선정을 **어떻게 받았는지** 한 자리에 적는다.

    조용히 다르게 동작하는 갈래를 두지 않는다(지시서 §2-3·§2-4·§2-5). 세 가지를
    말한다:

      · 손질에서 최불리를 안 눌렀다 → 도면 전체에서 뽑았다는 사실
      · 손질이 고른 것 중 전개가 **못 붙인** 헤드가 있다 → 개수와 목록
      · K 가 세션과 어긋나 손질 값을 썼다

    ★[두 화면 선정일치 §2-2] «채웠다» 를 말한다. 종전에는 `got["_filled"]`
      에 담기만 하고 **읽는 곳이 0곳**이었고, 이 함수의 문구는 그때에도
      「다른 헤드로 채우지 않았습니다」였다 — 채워 놓고 그렇게 적었다.
      프로그램이 옳은 일(K 유지)을 하면서 그 사실을 숨긴 것이 사용자가
      「위상이 깨졌다」로 읽은 이유다.

      `filled == 0` 이면 문구는 종전 그대로다(채우지 않았으므로 참이다).
    """
    note: dict = {"picked": len(picked), "k": k_use,
                  "candidates": got.get("candidate_heads"),
                  "from_edit": bool(picked)}
    msgs: list = []
    if not picked:
        msgs.append("손질에서 최불리를 정하지 않아 도면 전체에서 "
                    f"{k_use}개를 뽑았습니다.")
    else:
        cand = int(got.get("candidate_heads") or 0)
        lost = len(picked) - cand
        note["not_attachable"] = max(0, lost)
        if lost > 0:
            msgs.append(
                f"손질에서 고른 {len(picked)}개 중 {lost}개는 전개가 배관에 "
                f"붙이지 못했습니다 — 손질에서 그 헤드의 배관을 이어 주세요. "
                f"(다른 헤드로 채우지 않았습니다)")
        # ★[§2-2] 실제 교체를 **두 집합의 차**로 낸다 — 새로 세지 않는다.
        #   개수가 같은 채로 알맹이만 바뀔 수 있어(실측 30개 중 4개) 수만
        #   보고는 아무도 눈치채지 못한다.
        #   ★교체가 **없으면 칸 자체를 안 만든다.** 없는 일을 0 으로 적으면
        #     응답이 달라지고, 「교체 없는 세션은 종전과 한 바이트도 안
        #     다르다」는 기준(§4 기준 4)이 깨진다.
        fin = [int(i) for i in ((got.get("worst") or {}).get("heads") or ())]
        if fin:
            keep, want = set(fin), {int(i) for i in picked}
            out_rows = [_head_row(i, board, reasons)
                        for i in picked if int(i) not in keep]
            in_rows = [_head_row(i, board, None)
                       for i in fin if i not in want]
            if out_rows or in_rows:
                note["swapped_out"] = out_rows
                note["swapped_in"] = in_rows
        if filled > 0:
            note["filled"] = int(filled)
            # 못 붙는 헤드가 있었으므로 «후보 범위» 로 바꿔 채웠다. 그 수는
            # `cand`(후보 범위 ∩ 붙는 헤드)가 아니라 채운 수 그대로다.
            note["not_attachable"] = int(filled)
            # ★「기준개수를 지켰습니다」라고 **여기서 단정하지 않는다.** 이
            #   함수는 표가 서기 전에 돈다 — 제한 전개에서 또 떨어질 수 있어
            #   실제로 표가 K 미만으로 서는 일이 있다(실측 10 선정 → 표 9).
            #   단정해 놓으면 뒤이어 붙는 «9개 왔습니다» 와 서로 어긋난다.
            msgs.append(
                f"손질에서 고른 {len(picked)}개 중 {filled}개는 전개가 배관에 "
                f"붙이지 못해 같은 규칙(유하거리 긴 순서)으로 다음 순위 "
                f"{filled}개를 채웠습니다 — 선정은 {k_use}개를 채웠습니다. "
                f"빠진 자리는 도면에 표시했습니다.")
        if k_use != k_cfg:
            msgs.append(f"기준개수를 손질 값 {k_use}로 맞췄습니다 "
                        f"(설정에는 {k_cfg}이 남아 있었습니다).")
    note["messages"] = msgs
    for m in msgs:
        print(f"[최불리 인계] {m}")
    return note


def _handoff_after_table(got: dict, tbl, board) -> None:
    """[최불리 인계 §2-3] 손질이 고른 K개가 표에 다 왔는지 **표를 보고** 센다.

    물닿음 판정을 통과하고도 제한 전개에서 떨어질 수 있다. 그 자리를 좌표로
    되짚어 «어느 헤드가 빠졌는지» 까지 남긴다 — 개수만 세면 사람은 어디를
    고쳐야 할지 모른다.

    ★여기서 «채운다» 를 하지 않는다. 채우는 자리는 한 곳(`/design/build` 의
      백필)이고, 그때는 §2-2 가 무엇이 바뀌었는지 목록으로 말한다. 여기까지
      와서 수가 모자란 것은 **채우지 않은** 결손이므로 그렇게 적는다
      (S340 · D-F10-3).
    """
    note = got.get("handoff") or {}
    if not note.get("from_edit"):
        return
    picked = [int(i) for i in ((got.get("_picked") or ()))]
    if not picked:
        picked_n = int(note.get("picked") or 0)
    else:
        picked_n = len(picked)
    n_noz = len(getattr(tbl, "nozzles", None) or ())
    note["in_table"] = n_noz
    lost = picked_n - n_noz
    note["missing"] = max(0, lost)
    # ★[§2-2] **개수가 같아도** 알맹이가 바뀌었으면 지나가지 않는다.
    #
    #   종전에는 `lost <= 0` 하나로 막혔다. 그런데 다음 순위로 채우면 표의
    #   노즐 수는 K 그대로라 lost 가 0 이 되고, 「손질이 고른 30개 중 4개가
    #   빠지고 다른 4개가 들어온」 사실이 여기서 통째로 조용해졌다.
    swapped = int(note.get("filled") or 0)
    if lost <= 0 and not swapped:
        got["handoff"] = note
        return

    # 어느 헤드가 빠졌나 — 손질 disk 좌표를 표 좌표로 옮겨 맞대 본다.
    origin = got.get("origin_mm")
    disks = list(getattr(board, "disks", None) or ())
    miss: list = []
    if origin and disks and picked:
        import math as _m
        noz = {str(z.get("in")) for z in tbl.nozzles}
        at = [(float(n.get("x", 0) or 0), float(n.get("y", 0) or 0))
              for n in tbl.nodes if str(n.get("label")) in noz]
        # ★**1:1 로** 짝짓는다. 「각자 가장 가까운 것」만 보면 두 헤드가 같은
        #   노즐을 짚어도 둘 다 «찾음» 이 되어, 빠진 헤드를 놓친다.
        used: set = set()
        for i in picked:
            if i >= len(disks):
                continue
            q = (float(disks[i][0]) - float(origin[0]) + 1000.0,
                 float(disks[i][1]) - float(origin[1]) + 1000.0)
            best, bd = None, 1e18
            for t, p in enumerate(at):
                if t in used:
                    continue
                dd = _m.dist(p, q)
                if dd < bd:
                    best, bd = t, dd
            if best is None or bd > 100.0:
                miss.append({"disk": i,
                             "xy": [round(float(disks[i][0]), 1),
                                    round(float(disks[i][1]), 1)]})
            else:
                used.add(best)
    note["missing_heads"] = miss[:40]
    if lost <= 0:
        # 개수는 K 그대로다 — 교체가 있었을 뿐이고, 그 사실은 §2-2 가 이미
        # `swapped_out`/`swapped_in` 으로 말한다. 여기서는 **표를 보고 다시
        # 센 결과**만 적어 둔다. 두 수가 어긋나면 원인이 하나 더 있다는 뜻이라
        # 그때 이 값이 그것을 가리킨다(빠진 자리를 좌표로 되짚은 수).
        note["missing_by_table"] = len(miss)
        got["handoff"] = note
        return
    # ★이유를 **가른다.** 종전에는 전부 「배관이 끊겼다」로 적었는데, 실측에서
    #   그 헤드는 끊긴 것이 아니라 **다른 헤드와 같은 자리에 겹쳐** 있었다
    #   (대명동 (260307.6,−228066.2) · 전체망에서는 21개). 틀린 이유를 적으면
    #   사람을 엉뚱한 데로 보낸다 — 배관을 이으러 가도 고칠 것이 없다.
    n_share = len(got.get("shared_head_idx") or ())
    note["shared"] = n_share
    if swapped:
        # ★채워 놓고 「채우지 않았습니다」라고 적으면 안 된다. 채웠는데도
        #   모자란 경우다 — 영역 안에 채울 것이 더 없었다는 뜻이다(영역을
        #   넘지 않기 때문). 무엇이 모자란지를 그대로 말한다.
        msg = (f"손질에서 고른 {picked_n}개 중 {swapped}개가 배관에 안 붙어 "
               f"다음 순위로 채웠지만, 표에는 {n_noz}개만 왔습니다 — "
               f"기준개수 {picked_n}에 {lost}개 모자랍니다. "
               f"설계면적은 한 구역 안이라 다른 영역에서 끌어오지 않습니다 — "
               f"영역을 넓히거나 그 헤드의 배관을 이어 주세요.")
    elif n_share:
        msg = (f"손질에서 고른 {picked_n}개 중 {lost}개가 표에 오지 못했습니다 — "
               f"다른 헤드와 **같은 자리**라 하나로 합쳐졌습니다"
               f" (도면에 헤드 기호가 겹쳐 그려진 자리입니다)."
               f" 찍기에서 한쪽 묶음을 빼면 갈라집니다."
               f" (다른 헤드로 채우지 않았습니다)")
    else:
        msg = (f"손질에서 고른 {picked_n}개 중 {lost}개가 표에 오지 못했습니다 — "
               f"그 헤드의 배관이 전개에서 끊긴 자리입니다. 손질에서 이어 주세요. "
               f"(다른 헤드로 채우지 않았습니다)")
    note.setdefault("messages", []).append(msg)
    print(f"[최불리 인계] ★{msg}"
          + (f" · 자리 {[m['xy'] for m in miss[:6]]}" if miss else ""))
    got["handoff"] = note


def _classify_excluded(sess: dict, got: dict, board, probe=None) -> dict:
    """[F-5] 빠진 헤드를 네 갈래로 가른다 — «제외 2,864» 를 숫자로 쪼갠다.

        찍히지 않음     A 후보인데 board 에 없는 것 (suggest 를 돌린 세션만)
        이음 끊김       board 물길은 닿는데 전개가 못 붙인 것 (B4 부착 실패)
        물길 미도달     board 물길 자체가 안 닿는 것 (이음 끊김의 상류)
        고른 것 중 빠짐 손질이 골랐는데 이번 표에 못 들어간 것 (§2-3)

    좌표를 함께 돌려준다 — 화면이 분류별로 켜고 끌 수 있어야 어디를 이어야
    하는지 보인다. 숫자만 주면 «크다» 만 알고 «어디» 를 모른다.

    ★`probe` 를 받으면 **다시 재지 않는다.** 이 함수는 종전에 제 손으로
      `attachable_heads` 를 불렀는데 그것이 곧 **전체망 전개 한 번**이다
      (B1F 실측 117초). 같은 잡이 몇 줄 위에서 이미 쟀고(`attach.wet_heads`),
      게다가 여기 호출은 `selected_source` 를 안 넘겨 **다른 급수원 기준**의
      답이 나올 수 있었다 — 화면의 사유가 표와 다른 말을 하는 자리였다.
    """
    disks = getattr(board, "disks", []) or []
    total = len(disks)
    # board 물길 도달 — 급수원 성분에 붙은 헤드.
    try:
        wet_board = set((board.water_state() or {}).get("wet_heads") or [])
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 물길 분류 실패 — 미도달로 못 가른다: {exc}")
        wet_board = None
    # 전개가 붙일 수 있는 헤드 — 엔진의 공개 probe 를 그대로 쓴다.
    attach = None
    reasons: dict = {}
    try:
        if probe is None:
            from services.cad_import.design.restrict import attachable_heads
            es = sess.get("edit")
            probe = attachable_heads(es.convert_payload())
        if probe.get("ok"):
            attach = set(probe.get("wet") or ())
            reasons = dict(probe.get("reason") or {})
    except Exception as exc:  # noqa: BLE001
        print(f"[설계] 부착 probe 실패 — 이음 끊김을 못 가른다: {exc}")

    def xy(i):
        d = disks[i]
        return [round(float(d[0]), 1), round(float(d[1]), 1)]

    out = {"total": total}
    if wet_board is not None:
        dry = [i for i in range(total) if i not in wet_board]
        out["dry"] = {"n": len(dry), "xy": [xy(i) for i in dry]}
        if attach is not None:
            unatt = [i for i in wet_board if i not in attach and i < total]
            out["unattached"] = {"n": len(unatt), "xy": [xy(i) for i in unatt]}
            if reasons:
                # [§2-4] 「이음 끊김」 안에서도 갈래가 다르다 — 이을 것이 있는
                #   자리와 없는 자리(문양·스침)를 한 덩이로 세면 사람을 헛걸음
                #   시킨다. 수만 낸다(자리는 아래 «고른 것 중 빠짐» 이 짚는다).
                by: dict = {}
                for i in unatt:
                    w = reasons.get(i) or reasons.get(str(i))
                    if w:
                        by[str(w)] = by.get(str(w), 0) + 1
                if by:
                    out["unattached"]["why"] = by
    # ★[§2-3] «손질이 골랐는데 이번 표에 못 들어간 헤드» 를 따로 켠다.
    #
    #   `unattached` 는 도면 전체가 대상이라 수백 개가 될 수 있다(B1F 실측).
    #   사람이 볼 것은 «내가 고른 30개 중 빠진 4개» 다 — 그 넷이 수백 개
    #   사이에 묻히면 켜도 못 찾는다. 세는 자는 이미 `handoff` 가 갖고 있다.
    sw = [r for r in ((got.get("handoff") or {}).get("swapped_out") or ())
          if r.get("xy")]
    if sw:
        out["swapped_out"] = {"n": len(sw), "xy": [r["xy"] for r in sw],
                              "rows": sw}
    # 찍히지 않음 — suggest 후보 중 어느 board 헤드와도 250mm 안에 없는 것.
    cands = sess.get("suggest")
    if cands:
        import math as _m
        centers = [(float(d[0]), float(d[1])) for d in disks]
        missing = []
        for c_ in cands:
            cx, cy = float(c_["x"]), float(c_["y"])
            if not any(_m.hypot(cx - px, cy - py) <= 250.0
                       for px, py in centers):
                missing.append([round(cx, 1), round(cy, 1)])
        out["unpicked"] = {"n": len(missing), "xy": missing}
    return out
