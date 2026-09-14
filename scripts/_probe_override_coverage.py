# -*- coding: utf-8 -*-
"""[요소속성 수정카드 · 규칙 4] 고칠 수 있다고 **말해 놓고** 안 고치는 칸이 있나.

카드는 `FIELDS` 에 있는 것을 전부 「고칠 수 있습니다」로 보인다. 그런데 적용
자리는 넷으로 갈려 있고(사슬 전·표 전·표 직후·결합 직후), 어느 자리도 안 맡는
(갈래, 속성)이 하나라도 있으면 **사람이 값을 넣었는데 아무 일도 안 일어난다.**
그것이 못 옮김 목록에도 안 뜨면 규칙 4 위반이다 — 가장 알아채기 어려운 결함.

여기서는 `FIELDS` 를 **전수**로 돌려 셋 중 하나로 분류한다:

    적용됨      — 어딘가의 값이 실제로 바뀌었다
    못옮김보고  — 안 바뀌었지만 사유와 함께 목록에 올랐다 (정상)
    ★조용히    — 안 바뀌었고 목록에도 없다                (결함)

    python scripts/_probe_override_coverage.py
"""
from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from routes.module_f import overrides as ov                     # noqa: E402


class _Board:
    def __init__(self):
        self.pts = [(0.0, 0.0)] * 13
        self.disks = [[0.0, 0.0, 50.0]]
        self.hnodes = [{12}]


class _Tbl:
    def __init__(self, nodes, pipes):
        self.nodes, self.pipes = nodes, pipes


def _fixture():
    """배관 둘 · 세로 토막 둘 · 헤드 하나 — 네 갈래가 다 도는 최소 망."""
    nodes = {
        "N1": {"coords": [0.0, 0.0, 0.0], "type_id": "pump"},
        "N2": {"coords": [1.0, 0.0, 0.0], "type_id": "base"},
        "N3": {"coords": [2.0, 0.0, 0.0], "type_id": "head",
               "k_factor_si": 80.0, "required_pressure_bar": 1.0},
        "N4": {"coords": [2.0, 0.0, -0.3], "type_id": "base"},
        "N5": {"coords": [2.0, 0.0, -0.9], "type_id": "base"},
    }
    pipes = {
        "P1": {"start": "N1", "end": "N2", "length_m": 1.0, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
        "P2": {"start": "N2", "end": "N3", "length_m": 1.0, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
        "P3": {"start": "N3", "end": "N4", "length_m": 0.3, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
        "P4": {"start": "N4", "end": "N5", "length_m": 0.6, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
    }
    got = {"kfp": {"pipe_data": pipes, "nodes_meta_runtime": nodes},
           "edge_ref": {"P1": [10, 11], "P2": [11, 12]},
           "node_ref": {"N1": 10, "N2": 11, "N3": 12},
           "origin_mm": (1000.0, 1000.0)}
    tbl = _Tbl(
        nodes=[{"label": "1", "x": 0, "y": 0, "elevation": 0.0},
               {"label": "2", "x": 1000, "y": 0, "elevation": 0.0},
               {"label": "3", "x": 2000, "y": 0, "elevation": 0.0},
               {"label": "4", "x": 2000, "y": 0, "elevation": -0.3},
               {"label": "5", "x": 2000, "y": 0, "elevation": -0.9}],
        pipes=[{"label": p, "dia": 50, "dia_src": "규칙", "length": 1.0,
                "c": 120, "type": "Sch40"} for p in ("P1", "P2", "P3", "P4")])
    return got, _Board(), tbl


def _snap(got, tbl):
    return json.dumps({
        "kfp": got["kfp"],
        "nodes": tbl.nodes,
        "pipes": tbl.pipes,
    }, sort_keys=True, default=str)


# 갈래마다 «무엇을 가리킬까» — 위 표본에 실제로 있는 자리
def _key_for(kind):
    return {
        "pipe": ov.key_pipe(10, 11),           # P1
        "node": ov.key_node(11),               # N2 / 표 라벨 2
        "head": ov.key_head(0),                # N3 / 표 라벨 3
        "vert": ov.key_vert("head", 0, 2),     # P3 (아래 z −0.3)
    }.get(kind)


# 검증을 통과하는 «그럴듯한» 새 값
def _val_for(field):
    return {
        "length": 4.0, "dia": 65, "type": "Sch10", "c": 100.0,
        "roughness_mm": 0.05, "equivalent_length": 2.5,
        "elevation": -1.25, "k_factor_si": 115.2,
        "required_pressure_bar": 1.75,
    }[field]


def main() -> int:
    rows_out = []
    quiet = []

    # ── ①②③ 수리계산 갈래
    for (kind, field) in sorted(ov.FIELDS):
        if kind in ("sys", "mr"):
            continue
        got, board, tbl = _fixture()
        before = _snap(got, tbl)
        key = _key_for(kind)
        rows = ov.put([], key, kind, field, _val_for(field),
                      reason="전수 점검", at="t")
        try:
            _n1, m1, rep = ov.apply_to_kfp(got, board, rows)
            _n2, m2 = ov.apply_to_tables(tbl, got, board, rows, rep)
        except Exception as exc:                       # noqa: BLE001
            rows_out.append((kind, field, f"★예외 — {type(exc).__name__}: {exc}"))
            quiet.append((kind, field))
            continue
        changed = _snap(got, tbl) != before
        missed = list(m1) + list(m2)
        old_filled = rows[0].get("old") is not None
        if changed:
            verdict = "적용됨" + ("" if old_filled else " · ★원값 안 채움")
            if not old_filled:
                quiet.append((kind, field))
        elif missed:
            verdict = f"못옮김보고 — {missed[0].get('why')}"
        else:
            verdict = "★조용히 사라짐"
            quiet.append((kind, field))
        rows_out.append((kind, field, verdict))

    # ── ④ 통합 갈래
    import importlib.util as ilu
    spec = ilu.spec_from_file_location(
        "_mf_merge_fx", str(ROOT / "tests" / "test_module_f_merge.py"))
    fx = ilu.module_from_spec(spec)
    spec.loader.exec_module(fx)
    from routes.module_f.merge import merge_network

    base_merged = merge_network(fx._sample(), riser=fx._riser(),
                                mode="lsp_gravity")
    where = ov.merge_parts(base_merged)
    sys_pipe = next(p for p, k in where["pipe"].items() if k == "system")
    sys_node = next(n for n, k in where["node"].items() if k == "system")

    for (kind, field) in sorted(ov.FIELDS):
        if kind not in ("sys", "mr"):
            continue
        got = merge_network(fx._sample(), riser=fx._riser(), mode="lsp_gravity")
        # ★표본 입상관 행에는 `c` 칸이 아예 없다(`tests/test_module_f_merge.py`
        #   의 `_riser()`). 진짜 망에는 있다 — `emit_sdf` 가 `p["c"]` 를 반드시
        #   읽으므로, 없으면 결합 산출 자체가 KeyError 로 죽는다(실측).
        #   표본 때문에 「원값 없음」이 뜨는 것은 제품 결함이 아니므로,
        #   여기서 진짜 망의 모양으로 맞춰 두고 잰다.
        for _r in got["combined"].pipes:
            _r.setdefault("c", 120)
        lab = sys_node if field == "elevation" else sys_pipe
        before = json.dumps([dict(r) for r in got["combined"].pipes]
                            + [dict(n) for n in got["combined"].nodes],
                            sort_keys=True, default=str)
        rows = ov.put([], ov.key_merge(kind, lab), kind, field,
                      _val_for(field), reason="전수 점검", at="t")
        n, missed = ov.apply_to_merge(got, rows)
        after = json.dumps([dict(r) for r in got["combined"].pipes]
                           + [dict(x) for x in got["combined"].nodes],
                           sort_keys=True, default=str)
        changed = after != before
        old_filled = rows[0].get("old") is not None
        if kind == "mr":
            # 이 표본에는 기계실이 없다 — «못 옮김» 이 정상이다.
            verdict = ("적용됨(뜻밖)" if changed
                       else (f"못옮김보고 — {missed[0].get('why')}" if missed
                             else "★조용히 사라짐"))
            if not changed and not missed:
                quiet.append((kind, field))
        elif changed:
            verdict = "적용됨" + ("" if old_filled else " · ★원값 안 채움")
            if not old_filled:
                quiet.append((kind, field))
        elif missed:
            verdict = f"못옮김보고 — {missed[0].get('why')}"
        else:
            verdict = "★조용히 사라짐"
            quiet.append((kind, field))
        rows_out.append((kind, field, verdict))

    print("\n■ 고칠 수 있다고 말한 칸이 실제로 닿는가 — 전수\n")
    print(f"  {'갈래':6}{'속성':22}판정")
    print("  " + "─" * 68)
    for kind, field, verdict in rows_out:
        print(f"  {kind:6}{field:22}{verdict}")
    print(f"\n  FIELDS {len(ov.FIELDS)}개 · ★문제 {len(quiet)}개"
          + (f" — {quiet}" if quiet else " — 전부 닿거나 사유가 뜬다"))

    # ── 멱등: 같은 rows 로 두 번 적용해도 값도 원값도 그대로여야 한다
    print("\n■ 두 번 적용해도 같은가 (멱등 · 원값 보존)\n")
    bad = []
    for (kind, field) in sorted(ov.FIELDS):
        if kind in ("sys", "mr"):
            continue
        got, board, tbl = _fixture()
        rows = ov.put([], _key_for(kind), kind, field, _val_for(field),
                      reason="전수 점검", at="t")
        try:
            _a, _b, rep = ov.apply_to_kfp(got, board, rows)
            ov.apply_to_tables(tbl, got, board, rows, rep)
        except Exception:                              # noqa: BLE001
            continue
        once, old1 = _snap(got, tbl), copy.deepcopy(rows[0].get("old"))
        try:
            _a, _b, rep2 = ov.apply_to_kfp(got, board, rows)
            ov.apply_to_tables(tbl, got, board, rows, rep2)
        except Exception:                              # noqa: BLE001
            continue
        twice, old2 = _snap(got, tbl), rows[0].get("old")
        mark = []
        if once != twice:
            mark.append("★값이 또 바뀐다")
        if old1 != old2:
            mark.append(f"★원값이 덮인다 {old1} → {old2}")
        if mark:
            bad.append((kind, field, " · ".join(mark)))
        print(f"  {kind:6}{field:22}"
              + (" · ".join(mark) if mark else "그대로"))
    print(f"\n  ★멱등 위반 {len(bad)}개"
          + (f" — {[(k, f) for k, f, _ in bad]}" if bad else ""))

    n_bad = len(quiet) + len(bad)
    print("\n" + ("  ★전수 점검을 통과한다" if not n_bad
                  else f"  ★★{n_bad}건이 안 선다"))
    return 0 if not n_bad else 2


if __name__ == "__main__":
    raise SystemExit(main())
