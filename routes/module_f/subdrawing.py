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

# A 의 entity 는 색을 싣지 않는다(레이어만). 그래서 색은 **레이어 색**을 쓴다
# (`layer_colors`) — 파서가 DXF 에서 이미 읽어 둔 값이다. 캔버스가 레이어×색으로
# 묶으므로 결과는 여전히 «레이어 단위» 묶음이고, 화면에서는 도면 색 그대로 보인다.
# 텍스트 높이 — A 의 entity 에는 없어서 채워 넣는 자리다.
# ★종전 주석은 「0 이면 `extract_dia_text_points` 에서 걸린다」고 적었는데
#   **사실이 아니다** — 그 함수는 높이 칸을 받아서 버린다(`_lay,_col,x,y,_h,s`).
#   높이 0 을 실제로 거르는 곳은 `pipeline/stage45` 인데 계통도·기계실 경로는
#   그 단계를 타지 않는다. 즉 이 값이 «무엇을 막는가» 는 확인된 바 없다.
#   그래도 0 보다는 양수가 안전하다(높이를 쓰는 소비자가 생겨도 나눗셈·비교가
#   퇴화하지 않는다) — 근거를 실제 확인한 것으로만 줄여 적는다.
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


def layer_colors(parsed) -> dict:
    """레이어 이름 → DXF 색 번호(ACI). 파서가 이미 읽어 둔 것을 그대로 쓴다.

    ★꺼진 레이어는 색이 **음수**로 온다(CAD 의 관례). 계통도는 꺼둔 레이어에
      배관이 있는 일이 흔해서 우리는 그것도 읽는데(`include_hidden_layers`),
      음수를 그대로 넘기면 색표에서 못 찾아 전부 같은 색이 된다. 절댓값으로
      편다 — «안 보이게 해 둔 것» 과 «무슨 색인가» 는 다른 이야기다.
    """
    out: dict = {}
    for ly in ((parsed or {}).get("layers") or ()):
        try:
            c = int(ly.get("color", 7))
        except (TypeError, ValueError):
            c = 7
        out[str(ly.get("name"))] = abs(c) or 7
    return out


def entities_to_world(entities, colors=None) -> _EntWorld:
    """A 의 entity 목록 → 캔버스가 그릴 수 있는 World.

    폴리선은 마디마다 선분으로 편다 — 캔버스가 선분만 그리기 때문이고,
    계통도의 배관은 어차피 마디 단위로 잰다.

    `colors` : 레이어 → ACI 색 번호(`layer_colors`). 주면 **도면 색 그대로**
        그린다. 안 주면 종전처럼 한 색이다.

        ★종전에는 모든 도형에 색 7 을 박았다. 그래서 계통도·기계실이 통째로
          한 색으로 보였고, 배관·기호·건축선을 눈으로 가를 수가 없었다.
          평면도는 처음부터 도면 색으로 그려 왔다 — 두 화면이 다른 규칙을
          쓰고 있었던 셈이다.
    """
    w = _EntWorld()
    cmap = colors or {}
    for en in (entities or ()):
        t = en.get("t")
        lay = str(en.get("l") or "0")
        col = cmap.get(lay, 7)
        if t == "L":
            p = en.get("p") or []
            if len(p) >= 4:
                w.segs.append((lay, col, (float(p[0]), float(p[1])),
                               (float(p[2]), float(p[3]))))
        elif t == "PL":
            pts = en.get("p") or []
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                if len(a) >= 2 and len(b) >= 2:
                    w.segs.append((lay, col, (float(a[0]), float(a[1])),
                                   (float(b[0]), float(b[1]))))
        elif t == "C":
            c = en.get("c") or []
            if len(c) >= 2:
                w.circles.append((lay, col, float(c[0]), float(c[1]),
                                  float(en.get("r") or 0.0)))
        elif t == "A":
            c = en.get("c") or []
            if len(c) >= 2:
                w.arcs.append((lay, col, float(c[0]), float(c[1]),
                               float(en.get("r") or 0.0)))
                a = en.get("a") or [0.0, 360.0]
                sa = float(a[0])
                ea = float(a[1]) if len(a) > 1 else sa + 360.0
                sweep = (ea - sa) % 360.0 or 360.0
                w.arc_ang.append((sa, sweep))
        elif t == "T":
            p = en.get("p") or []
            if len(p) >= 2:
                w.texts.append((lay, col, float(p[0]), float(p[1]),
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


def layer_options(entities, colors=None) -> list[dict]:
    """이 도면의 레이어 목록 + A 의 이름 사전 분류.

    계통도·기계실도 «어느 선이 배관인가» 가 갈림길이다. 이름 사전은 추천일
    뿐이라(평면도에서 절반이 OTHER 로 떨어진다) 결정은 사람이 한다 — 여기서는
    고를 거리를 만들어 줄 뿐이다.
    """
    from routes.module_f.common import _layer_category
    n_by: dict[str, int] = {}
    for en in (entities or ()):
        nm = str(en.get("l") or "0")
        n_by[nm] = n_by.get(nm, 0) + 1
    # [BLOCKED §27] 이름만 보면 «정반대로» 읽는 자리가 있다 — 대명동 계통도의
    #   `SP` 는 사전이 배관으로 읽지만 헤드 기호 918개다. 여기는 entity 를
    #   갖고 있으므로 기하 교정을 적용한 표를 쓴다. 못 불러오면 이름만으로
    #   떨어진다 — 사람이 고르는 화면이 죽는 것보다는 낫다.
    cats = {}
    try:
        from remote30_prototype import categorize_layers
        cats = categorize_layers(entities)
    except Exception:  # noqa: BLE001
        cats = {}
    # 색도 함께 — 「색깔별로 보이게」 한 화면과 같은 색이라야 사람이 목록과
    # 도면을 맞대 볼 수 있다. 색표는 캔버스가 쓰는 그것 하나뿐이다.
    lc = colors or {}
    try:
        from services.cad_import.colors import rgb_dark
    except Exception:  # noqa: BLE001
        def rgb_dark(_c):
            return "#c8a064"
    out = []
    for nm, n in sorted(n_by.items(), key=lambda kv: (-kv[1], kv[0])):
        row = {"layer": nm, "n": n,
               "cat": cats.get(nm) or _layer_category(nm)}
        if nm in lc:
            row["color"] = lc[nm]
            row["css"] = rgb_dark(lc[nm])
        out.append(row)
    return out


def path_graph(entities, *, layer_filter=None):
    """경로 그래프 한 벌 — **추출이 쓰는 바로 그것** 을 그대로 만든다.

    ★미리보기와 추출이 다른 그래프를 쓰면 화면이 거짓말을 한다. 그래서 여기서
      A 의 `build_system_graph` 를 추출과 **같은 인자** 로 부른다
      (`force_connect=True` — 계통도/기계실 전용, 평면도 경로에서는 금지).

    반환: (graph, edge_len, stats) — A 의 것 그대로.
    """
    from remote30_prototype import build_system_graph
    return build_system_graph(entities, layer_filter=layer_filter,
                              force_connect=True)


def pick_system_layer(entities, a_xy, b_xy, *, snap_tolerance_mm=2500.0):
    """찍은 두 점이 **어느 계통(레이어)** 에 있는지 골라 준다.

    ★이것이 「경로가 꼬인다」의 알맹이다. 필터를 안 걸면 엔진의 이름 사전이
      배관 레이어를 **여러 장 한꺼번에** 고른다. 대명동 계통도는 HSP(고층)·
      LSP(저층)·감압밸브가 그렇게 한 그래프에 섞이는데, 둘은 서로 다른
      배관이라 최단경로가 그 사이를 넘나든다. 실측한 경로가 그랬다:

        자동(섞음) — LSP → HSP → LSP · 넘나듦 4회 · 연장 126.3 m
        HSP 만     — HSP            · 넘나듦 0회 · 연장 118.0 m

      (기계실도 같다: -소화(SP-저) → -소화(SP-고) 로 2회 넘나든다.)

      추정 이음(허용오차 다리)은 범인이 아니었다 — 경로 126.3 m 중 3곳 1.0 m
      (1%)뿐이고, 그것을 뒤로 미뤄도 경로가 그대로였다. 그래서 다리를 손대는
      대신 **계통을 하나로 좁힌다.**

    고르는 규칙: **두 점이 각각 가장 붙어 있는** 레이어가 같고, 그것이
    2등보다 뚜렷하게 가까울 때만 그 레이어로 좁힌다. 아니면 고르지 않는다
    (None) — 억지로 좁히면 엉뚱한 계통 안에서 먼 길을 돌게 된다.

    ★«허용거리 안이면 된다» 로 걸면 안 된다. 추출(`extract_system`)은 클릭을
      거리 제한 없이 가장 가까운 절점에 붙인다(사용자 요구로 그렇게 만들었다).
      그래서 화면에서 사람이 찍는 자리는 배관에서 수십 m 떨어져 있는 것이
      보통이다 — 실측: 계통도에서 632 m. 절대 허용거리로 후보를 자르면 좁히기가
      **한 번도 안 걸린다.** 자는 «어느 계통에 더 붙었나» 라는 **견줌**이다.

    ★«경로가 짧은 쪽» 만 보면 또 틀린다. 두 계통이 나란히 붙어 있으면 찍은 점
      위의 계통(HSP·13.0 m)을 두고 옆 계통(LSP·9.0 m)을 고른다 — 시험에서
      실제로 그랬다. 사람이 찍은 자리가 먼저고, 길이는 동점일 때만 본다.

    판단 근거는 함께 돌려주어 화면이 말할 수 있게 한다(조용히 바꾸지 않는다).
    """
    import math

    from remote30_prototype import (_nearest_graph_node, _shortest_path,
                                    build_system_graph)
    _g, _el, stats = build_system_graph(entities, layer_filter=None,
                                        force_connect=False)
    cands = sorted(stats.get("layer_filter_used") or ())
    if len(cands) < 2:
        return None, {"candidates": [], "reason": "섞인 레이어가 없습니다"}
    rows = []
    for nm in cands:
        try:
            g, el, st = build_system_graph(entities, layer_filter={nm},
                                           force_connect=True)
        except Exception:  # noqa: BLE001
            continue
        if not g:
            continue
        row = {"layer": nm, "nodes": len(g)}
        ds, ns = [], []
        for xy in (a_xy, b_xy):
            n = _nearest_graph_node(g, (float(xy[0]), float(xy[1])))
            if n is None:
                ds.append(float("inf"))
                ns.append(None)
            else:
                ds.append(math.hypot(n[0] - float(xy[0]), n[1] - float(xy[1])))
                ns.append(n)
        row["snap_a_mm"], row["snap_b_mm"] = (round(d, 1) for d in ds)
        row["snap_mm"] = round(max(ds), 1)
        row["ok"] = bool(ns[0] and ns[1])
        if row["ok"]:
            fk = set()
            for (ea, eb) in (st.get("forced_bridge_edges") or ()):
                ka, kb = (int(ea[0]), int(ea[1])), (int(eb[0]), int(eb[1]))
                fk.add((min(ka, kb), max(ka, kb)))
            p = _shortest_path(g, el, ns[0], ns[1], penalty_keys=fk)
            if p and len(p) >= 2:
                tot = 0.0
                for u, v in zip(p, p[1:]):
                    ln = el.get((min(u, v), max(u, v)))
                    tot += ln if ln is not None else math.hypot(
                        u[0] - v[0], u[1] - v[1])
                row["path_m"] = round(tot / 1000.0, 1)
                row["_len"] = tot
            else:
                row["ok"] = False
                row["why"] = "그 레이어 안에서는 두 점이 안 이어집니다"
        rows.append(row)

    live = [r for r in rows if r["ok"]]
    diag = {"candidates": rows, "snap_tolerance_mm": float(snap_tolerance_mm)}
    if not live:
        diag["reason"] = "두 점이 이어지는 레이어가 없습니다 — 섞은 채로 뽑습니다"
        return None, diag
    # 동점의 자는 **도면의 축척을 따른다.** 엔진이 «같은 점» 으로 보는 거리
    # (`snap_eps_mm` = SNAP_TOL_MM × 축척)가 그 자다. 고정 mm 로 걸면 축척이
    # 다른 도면에서 무너진다 — 용지축척 계통도는 0.259mm, 실좌표 평면도는
    # 17.4mm 로 60배 넘게 벌어진다. 고정 250mm 를 쓰면 스키매틱 도면에서는
    # 모든 레이어가 «동점» 이 되어 좁히기가 한 번도 안 걸린다(실측으로 그랬다).
    tie = float(stats.get("snap_eps_mm") or 0.0) or 1.0
    diag["tie_mm"] = round(tie, 3)

    def _near(key):
        ranked = sorted(live, key=lambda r: (r[key], r["_len"]))
        top = ranked[0]
        rest = [r for r in ranked[1:] if r[key] > top[key] + tie]
        return top, (len(rest) == len(ranked) - 1)

    ta, clear_a = _near("snap_a_mm")
    tb, clear_b = _near("snap_b_mm")
    if ta["layer"] != tb["layer"]:
        diag["reason"] = (f"찍은 두 점이 서로 다른 계통에 붙습니다"
                          f"(«{ta['layer']}» · «{tb['layer']}») — "
                          f"섞은 채로 뽑습니다")
        return None, diag
    if not (clear_a and clear_b):
        diag["reason"] = ("어느 계통에 붙은 것인지 가릴 만큼 차이가 없습니다 — "
                          "섞은 채로 뽑습니다")
        return None, diag
    diag["reason"] = f"찍은 두 점이 «{ta['layer']}» 배관에 가장 붙어 있습니다"
    for r in rows:
        r.pop("_len", None)
    return ta["layer"], diag


def graph_payload(entities, *, layer_filter=None) -> dict:
    """경로 그래프를 화면이 읽을 모양으로.

    실측(계통도·기계실 4장): 노드 132~382 · 간선 131~411 · JSON 3~8KB.
    이 정도면 통째로 내려보내고 브라우저가 직접 최단경로를 풀 수 있다 —
    마우스가 움직일 때마다 서버를 왕복하면 LAN·터널에서 눈에 띄게 밀린다.

    ★추측 연결(force_connect 가 이은 직선)은 «따로» 실어 보낸다. 실측 배관과
      한 모양으로 그리면 사람이 확인한 것과 기계가 고른 것을 구별할 수 없다.
    """
    graph, edge_len, stats = path_graph(entities, layer_filter=layer_filter)
    # ★자동으로 «배관» 이라고 고른 레이어들. 계통도는 고층·저층 입상관이
    #   서로 다른 레이어에 있어서(HSP·LSP·감압밸브), 그 셋을 한 그래프에
    #   섞으면 경로가 계통 사이를 오갈 수 있다 — 사람 눈에는 «꼬인» 길이다.
    #   화면이 그 사실을 말할 수 있게 실어 보낸다.
    auto_layers = sorted(stats.get("layer_filter_used") or ())
    idx = {n: i for i, n in enumerate(graph)}
    forced = set()
    for (ea, eb) in (stats.get("forced_bridge_edges") or []):
        ka, kb = idx.get(tuple(ea)), idx.get(tuple(eb))
        if ka is not None and kb is not None:
            forced.add((min(ka, kb), max(ka, kb)))
    seen, edges = set(), []
    for a, nbrs in graph.items():
        for bnode in nbrs:
            ia, ib = idx[a], idx[bnode]
            key = (min(ia, ib), max(ia, ib))
            if key in seen:
                continue
            seen.add(key)
            ln = edge_len.get((a, bnode)) or edge_len.get((bnode, a)) or 0.0
            edges.append([key[0], key[1], round(float(ln), 1),
                          1 if key in forced else 0])
    return {
        "nodes": [[int(round(n[0])), int(round(n[1]))] for n in graph],
        "edges": edges,
        "auto_layers": auto_layers,
        "components": stats.get("components_after_bridge"),
        "bridges": stats.get("bridges_applied"),
        "forced": len(forced),
        # ★추측연결 벌점을 **숫자로 실어 보낸다.** 화면이 제 값을 적으면
        #   미리보기와 추출이 다른 길을 고를 수 있다 — 실제로 그랬다: 서버는
        #   1e9, 화면은 1e6 으로 1000배 달랐다. 도면 단위가 mm 라 1e6 은 1km 이고,
        #   큰 도면에서는 실배관 우회가 그 값에 닿을 수 있다. 그러면 화면은
        #   추측 직선을, 서버는 실배관을 고른다 — 미리보기가 거짓말이 된다.
        #   (이 그래프로 뽑는 `extract_system`·`extract_machineroom` 이
        #    `_shortest_path` 를 그 기본값으로 부른다.)
        "forced_penalty_mm": _forced_penalty_mm(),
    }


def _forced_penalty_mm() -> float:
    """추출이 실제로 쓰는 벌점 — 엔진 기본값을 그대로 읽는다.

    여기서 숫자를 다시 적지 않는다. 엔진이 값을 바꾸면 화면도 따라가야 하고,
    그러려면 «적힌 곳» 이 하나여야 한다.
    """
    import inspect

    from remote30_graph import _shortest_path
    p = inspect.signature(_shortest_path).parameters.get("penalty_mm")
    if p is not None and p.default is not inspect.Parameter.empty:
        return float(p.default)
    # 서명이 바뀌면 조용히 다른 값을 쓰지 않는다 — 알 수 없으면 말한다.
    raise RuntimeError(
        "추측연결 벌점을 엔진에서 읽지 못했습니다 — `_shortest_path` 의 "
        "`penalty_mm` 기본값이 사라졌습니다. 화면 미리보기가 추출과 다른 길을 "
        "고를 수 있으므로 여기서 임의 값을 쓰지 않습니다.")


def extract_system(entities, pump_xy, av_xy, *, snap_tolerance_mm=2500.0,
                   waypoints=None, floor_profile_rows=None,
                   layer_filter=None) -> dict:
    """S720 — 계통도에서 펌프 → 알람밸브 경로(입상관)를 뽑는다.

    실패(클릭이 배관에서 너무 멀다 · 두 점이 안 이어진다)는 `ValueError` 로
    올라온다. 임의로 메우지 않는다 — 특허 S340 의 «실제 배관만으로 이어지지
    아니하는 자리는 임의로 메우지 아니하고 미도달로 보고한다» 를 따른다.
    """
    from remote30_prototype import extract_system_path
    return extract_system_path(
        entities, tuple(pump_xy), tuple(av_xy),
        snap_tolerance_mm=float(snap_tolerance_mm),
        layer_filter=layer_filter or None,
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
                        snap_tolerance_mm=2500.0, ceiling_m=None,
                        layer_filter=None) -> dict:
    """S730 — 기계실에서 수원(탱크) → 입상관 연결점 경로를 뽑는다.

    좌표는 **평면 그대로 보존**된다(H-D6). 계통도처럼 세로 막대로 재배치하면
    기계실 배관의 평면 형상이 뭉개진다 — A 도 그래서 둘을 갈라 놓았다.

    라벨은 `m1..mK` 라 라이저(1~10)·헤드(10+)와 부딪히지 않는다.
    """
    from remote30_prototype import extract_machine_room_path
    return extract_machine_room_path(
        entities, tuple(source_xy), tuple(conn_xy),
        snap_tolerance_mm=float(snap_tolerance_mm),
        layer_filter=layer_filter or None,
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
