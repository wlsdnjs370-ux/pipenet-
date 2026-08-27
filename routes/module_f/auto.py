# -*- coding: utf-8 -*-
"""평면도 «자동 추출» — 모듈 A 의 위상 검출 경로.

같은 평면도에서 같은 것(최불리 헤드군)을 뽑는 길이 둘이다. 순서가 다르다.

    A · 자동   도면 → **알람밸브 한 점** + **헤드 영역** 을 사람이 정하고,
               헤드 검출 · 그래프 복원 · 앵커 · 최불리 K 는 전부 자동
    E · 수동   도면 → 사람이 **색(레이어×색) 으로 배관을 찍고** 헤드를 찍고
               급수원을 찍은 뒤, 그 위에서 손질하고 최불리를 고른다

자동이 빠르고, 수동이 지저분한 도면에서 버틴다. 어느 쪽이 옳다가 아니라
**도면이 정하는 문제**라, 고르는 것은 사람이다(업로드 때 고른다).

여기서 새 알고리즘은 쓰지 않는다. A 의 `select_worst30_heads_anchored` 와
`build_input_tables` 를 그대로 부른다 — A 는 읽기 전용이다.

★A 의 표는 **알람밸브가 라벨 10** 이다(`build_input_tables` 의 `counter = [10]`).
  G 의 표는 BFS 로 1부터 매기므로 통합(S740) 앞에서 +9 를 먹였는데, A 의 표는
  이미 그 규약이라 **옮기면 안 된다**. 두 경로의 표가 같은 자리에 들어가므로
  이 차이를 여기서 못박아 둔다(`LABEL_OFFSET_FOR_AUTO = 0`).
"""
from __future__ import annotations

# A 의 표는 이미 기준점이 10 이다 — 통합 앞에서 라벨을 옮기지 않는다.
LABEL_OFFSET_FOR_AUTO = 0


class AutoError(ValueError):
    """자동 추출이 성립하지 않는다 — 임의로 메우지 않고 올린다."""


def parse_plan(dxf_path):
    """평면도 한 장 → (entities, layer_categories, 진단).

    A 의 `parse_dxf_bundle` 을 쓴다 — 헤드 검출·그래프 복원이 이 파서의 entity
    모양을 전제한다. 시각화용 파서(`parse_dxf_for_view`)와 섞으면 레이어 승격
    (헤드 틈 지문으로 PIPE 로 올리는 것) 같은 판정이 통째로 빠진다.

    ★캐시본을 쓴다. 화면은 도면을 이미 찍기판으로 한 번 읽었고(그쪽이 싸다),
      자동을 고르면 A 의 파서로 **또** 읽어야 한다 — 그 값이 실측으로 크다
      (LH306 16MB 에서 5.0s, 큰 도면은 수십 초). A 는 파일 «내용 해시» 로
      디스크 캐시하는 `parse_dxf_bundle_cached` 를 이미 갖고 있는데 안 쓰고
      있었다: 같은 도면을 다시 열면 5.02s → 0.06s (86배).

      내용 해시 키라 파일을 다시 올려도(mtime 만 바뀌어도) 캐시가 산다 —
      handoff 가 mtime 때문에 캐시를 통째로 버리던 것과 같은 함정을 여기서는
      처음부터 피한다.
    """
    from remote30_prototype import _categorize_layer, parse_dxf_bundle_cached

    bundle = parse_dxf_bundle_cached(dxf_path)
    ents = list(bundle.entities or ())
    if not ents:
        raise AutoError("도면에서 도형을 읽지 못했습니다.")

    names = {str(e.get("l") or "0") for e in ents}
    layer_cat = {}
    for n in names:
        try:
            layer_cat[n] = _categorize_layer(n)
        except Exception:  # noqa: BLE001 — 한 이름이 막혀도 나머지는 분류한다
            layer_cat[n] = "OTHER"

    diag = {
        "entities": len(ents),
        "layers": len(names),
        "cats": _cat_counts(layer_cat),
        # 외부참조 시트면 도면 내용이 딴 파일에 있다 — 헤드 0개로 끝난다.
        "xref": dict(bundle.xref_diagnostics or {}),
        "promoted": list(bundle.promoted_layers or ()),
        "bbox": list(bundle.bbox or ()),
    }
    return ents, layer_cat, diag


def _cat_counts(layer_cat: dict) -> dict:
    out: dict[str, int] = {}
    for c in layer_cat.values():
        out[c] = out.get(c, 0) + 1
    return out


def detect_head_candidates(entities, layer_cat, rects=None):
    """헤드 후보만 먼저 — 영역을 정하기 전에 «어디에 헤드가 있나» 를 본다."""
    from remote30_prototype import detect_heads

    # HeadDetection(pos · bbox · kind · confidence · block_name · layer)
    heads = detect_heads(entities, layer_cat)
    out = []
    for h in heads or ():
        pos = getattr(h, "pos", None)
        if not pos:
            continue
        out.append({"x": float(pos[0]), "y": float(pos[1]),
                    "conf": float(getattr(h, "confidence", 0.0) or 0.0),
                    "kind": str(getattr(h, "kind", "") or ""),
                    "why": str(getattr(h, "block_name", "") or "")})
    if rects:
        out = [h for h in out
               if any(x0 <= h["x"] <= x1 and y0 <= h["y"] <= y1
                      for x0, y0, x1, y1 in rects)]
    return out


# 검출한 헤드를 감싸는 사각형에 둘 여유(mm). 헤드 중심만으로 자르면 기호 반경과
# 접속관 끝이 경계 밖으로 나가 region 게이트에서 떨어진다.
AUTO_REGION_PAD_MM = 1000.0


def head_region_of(rects):
    """사각형 목록 → A 의 `HeadRegion`."""
    if not rects:
        raise AutoError("영역이 비었습니다.")
    from remote30_graph import HeadRegion
    return HeadRegion.from_rects([tuple(float(v) for v in r) for r in rects])


def region_around(heads, pad_mm: float = AUTO_REGION_PAD_MM):
    """검출한 헤드 전체를 감싸는 사각형 하나.

    `select_worst30_heads_anchored` 는 `head_region` 을 필수로 받는다 — 앵커를
    세우려면 «어디의 헤드인가» 가 있어야 하기 때문이다. 그런데 사람 입장에서
    기본값은 «도면에서 찾은 헤드 전부» 다. 영역을 그리는 것은 그것을 «좁히는»
    선택이지, 시작하기 위한 조건이 아니다.

    그래서 안 그렸으면 검출 결과에서 만들어 쓴다.
    """
    pts = [(float(h["x"]), float(h["y"])) for h in (heads or ())]
    if not pts:
        raise AutoError(
            "도면에서 헤드를 찾지 못했습니다 — 헤드 검출을 먼저 해 보고, "
            "그래도 0개면 이 도면은 자동 추출로 읽을 수 없습니다.")
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return [[min(xs) - pad_mm, min(ys) - pad_mm,
             max(xs) + pad_mm, max(ys) + pad_mm]]


def run_auto(entities, layer_cat, *, alarm_xy, rects, k: int = 30,
             project_title: str = "모듈 F 자동 추출", progress_cb=None):
    """A 의 anchored 선정 → 5종 입력표.

    반환: {"tables", "selection", "summary"} — `tables` 는 A 의 `PipeTables`,
    필드가 G 의 `PipeTablesG` 와 같아 하류(수리계산 표·통합·산출)가 그대로 받는다.
    """
    from remote30_prototype import (build_input_tables,
                                    select_worst30_heads_anchored)

    if alarm_xy is None:
        raise AutoError("알람밸브 위치를 도면에서 찍으세요.")
    # 영역은 «좁히는» 선택이다 — 안 그렸으면 검출한 헤드 전부를 범위로 삼는다.
    region_auto = not rects
    if region_auto:
        rects = region_around(detect_head_candidates(entities, layer_cat))
    region = head_region_of(rects)
    audit: dict = {}
    try:
        sel = select_worst30_heads_anchored(
            entities, layer_cat, (float(alarm_xy[0]), float(alarm_xy[1])),
            region, k=int(k), audit_out=audit, progress_cb=progress_cb)
    except ValueError as exc:
        raise AutoError(str(exc)) from None

    if not sel.heads:
        raise AutoError(
            "영역 안에서 급수원에 닿는 헤드를 찾지 못했습니다 — 알람밸브 위치나 "
            "영역을 다시 잡아 보세요.")

    tables = build_input_tables(
        sel, entities, project_title=project_title,
        anchor_window=audit.get("anchor_window"))

    summary = summarize(sel, tables)
    # 영역을 사람이 그린 것인지 검출에서 만든 것인지 — 화면이 말할 수 있게.
    summary["region_auto"] = region_auto
    summary["zones"] = len(rects)
    return {"tables": tables, "selection": sel, "rects": rects,
            "summary": summary}


def summarize(sel, tables) -> dict:
    """자동 추출 결과 한 장 — 수동 경로의 요약과 같은 것을 말한다."""
    dists = [float(d) for d in (getattr(sel, "distances", None) or ())]
    meta = dict(getattr(tables, "meta", None) or ())
    return {
        "k": len(getattr(sel, "heads", None) or ()),
        "far_m": round(max(dists) / 1000.0, 2) if dists else 0.0,
        "near_m": round(min(dists) / 1000.0, 2) if dists else 0.0,
        "edges": len(getattr(sel, "edges", None) or ()),
        "nodes": len(getattr(tables, "nodes", None) or ()),
        "pipes": len(getattr(tables, "pipes", None) or ()),
        "nozzles": len(getattr(tables, "nozzles", None) or ()),
        "fittings": len(getattr(tables, "fittings", None) or ()),
        # 급수원이 그래프에서 떨어져 있었나 — 0 이 아니면 이어 붙인 것이다.
        "source_bridge_mm": round(
            float(getattr(sel, "source_bridge_dist_mm", 0.0) or 0.0), 1),
        "source_fallback": bool(getattr(sel, "source_fallback", False)),
        "anchor_label": meta.get("앵커 노드"),
    }


def preview_view(tables) -> dict:
    """캔버스가 그릴 수 있는 최소 payload — 노드·배관·헤드.

    수동 경로의 설계 미리보기와 좌표계가 같다(표의 x·y 는 mm).
    """
    nodes = []
    heads = {str(r.get("in")) for r in (getattr(tables, "nozzles", None) or ())}
    for n in (getattr(tables, "nodes", None) or ()):
        lab = str(n.get("label"))
        rec = {"label": lab, "x": float(n.get("x", 0) or 0),
               "y": float(n.get("y", 0) or 0)}
        if lab in heads:
            rec["head"] = True
        if str(n.get("io_node")) == "Input":
            rec["input"] = True
        nodes.append(rec)
    pipes = [{"label": str(r.get("label")), "a": str(r.get("in")),
              "b": str(r.get("out")), "dia": r.get("dia"),
              "len_m": r.get("length")}
             for r in (getattr(tables, "pipes", None) or ())]
    return {"nodes": nodes, "pipes": pipes}
