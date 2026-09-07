# -*- coding: utf-8 -*-
"""[§27 확인] 두 점 클릭으로 «따라오는» 배관망이 지금 실제로 서는가.

§27 은 두 갈래 증상을 적었다 — ①헤드 인식 ②경로 추출. 이번에 넣은 것은
«이 묶음만 크게»(보기 도우미)뿐이므로, 경로 추출 자체가 나아졌는지는 **별개**다.

★처음에 「조각 1개 · 100%」로 재고 «건강하다» 고 읽을 뻔했다. 그건 만들어진
  값이다 — `force_connect=True` 가 남은 조각을 **거리 무제한**으로 이어 붙이므로
  마지막 조각 수는 언제나 1 이다. 볼 것은 **이어 붙이기 전** 이다:

    · 조각(강제연결 전)  — 진짜 배관이 몇 덩이로 끊겨 있나
    · 추측연결(forced)   — 그 중 몇 군데를 «없는 배관» 으로 메웠나
    · 가장 큰 덩이 %     — 두 점이 같은 덩이에 있을 확률

  추측연결을 밟는 경로는 도면에 없는 직선을 따라간다. 그래서 이 수가 곧
  「따라오는 선이 얼마나 지어낸 것인가」다.

★`layer_filter=None` 은 «도면 전체» 가 아니다 — **이름 사전이 고른 레이어** 다
  (`_auto_pipe_layer_filter`). §27 이 「정반대로 읽는다」고 지목한 바로 그것이다.
"""
from __future__ import annotations

import glob
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
sys.path.insert(0, _ROOT)

from routes.module_f.common import _boot          # noqa: E402

_boot()

from routes.module_f.subdrawing import layer_options   # noqa: E402


def _components(graph):
    seen, out = set(), []
    for start in graph:
        if start in seen:
            continue
        stack, n = [start], 0
        seen.add(start)
        while stack:
            cur = stack.pop()
            n += 1
            for nb in graph[cur]:
                if nb not in seen:
                    seen.add(nb)
                    stack.append(nb)
        out.append(n)
    return sorted(out, reverse=True)


def _measure(ents, layers):
    """강제 연결 **전** 상태까지 함께 본다 — 뒤만 보면 늘 «건강» 하다."""
    from remote30_prototype import build_system_graph

    g0, _e0, s0 = build_system_graph(ents, layer_filter=layers,
                                     force_connect=False)
    comps = _components(g0)
    big = comps[0] if comps else 0
    _g1, _e1, s1 = build_system_graph(ents, layer_filter=layers,
                                      force_connect=True)
    return {
        "nodes": s0["node_count"],
        "lines": s0["line_entity_count"],
        "comps": len(comps),
        "big_pct": big / max(len(g0), 1) * 100,
        "forced": s1["forced_bridges"],
        "used": s0.get("layer_filter_used"),
        "fallback": s0.get("layer_filter_fallback_no_match"),
    }


def _row(tag, m):
    print(f"  {tag:<24} 노드 {m['nodes']:>5} · 선 {m['lines']:>5}"
          f" · 조각(강제 전) {m['comps']:>4}"
          f" · 최대덩이 {m['big_pct']:>5.1f}%"
          f" · 추측연결 {m['forced']:>4}")


def main() -> int:
    names = sys.argv[1:] or [
        "1. 입력도면 대명동 단위세대 계통도.dxf",
        "1. 입력도면 대명동 단위세대 기계실.dxf",
    ]
    from routes.module_f.subdrawing import parse_subdrawing

    for name in names:
        hits = glob.glob(os.path.join(_ROOT, "**", name), recursive=True)
        if not hits:
            print(f"\n[{name}] 파일 없음 — 건너뜀")
            continue
        print(f"\n[{os.path.basename(hits[0])}]")
        ents, _ = parse_subdrawing(hits[0])
        auto = _measure(ents, None)
        print(f"  이름 사전이 고른 레이어: {auto['used']}"
              + ("  ★못 골라 전체로 폴백" if auto["fallback"] else ""))
        _row("사전이 고른 대로(기본)", auto)

        opts = layer_options(ents)
        cand = []
        for o in opts:
            m = _measure(ents, {o["layer"]})
            if m["nodes"] >= 20:
                cand.append((m["nodes"], o["layer"], m))
        cand.sort(reverse=True)
        for _n, lname, m in cand[:5]:
            _row(f"사람이 «{lname}» 만 고름", m)
        if len(cand) >= 2:
            pick = {lname for _n, lname, _m in cand[:2]}
            _row(f"«{'+'.join(sorted(pick))}» 함께", _measure(ents, pick))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
