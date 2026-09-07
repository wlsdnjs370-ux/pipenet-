# -*- coding: utf-8 -*-
"""[§27-1] 이름 사전이 «정반대로» 읽는 자리를 기하로 뒤집을 수 있나 — 교정.

§27 은 묶음(bundle) 단위 지표로는 못 가른다고 두 번 확인했다(선분 길이·접점).
그러나 그때 같이 적어 둔 문장이 길을 남긴다:

    「entity 단위(닫힘·도형 지름)로는 갈리지만 그 정보는 묶음에 없다」

`_categorize_layer` 는 **이름만** 받는다(`name: str`). 부르는 쪽은 entity 를
갖고 있다. 그러니 여기서는 그 정보가 있다.

가설 — **PIPE 로 읽힌 레이어의 내용이 압도적으로 «작고 닫힌 도형» 이면 그것은
기호(헤드) 레이어지 배관이 아니다.**

붙이기 전에 잰다. 이름 사전이 «맞는» 도면을 이 규칙이 망가뜨리면 안 된다.
"""
from __future__ import annotations

import glob
import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
sys.path.insert(0, _ROOT)

from routes.module_f.common import _boot          # noqa: E402

_boot()

import remote30_prototype as RP                   # noqa: E402

CLOSE_EPS_MM = 1.0          # 첫점≈끝점 판정


def _stats(ents):
    """레이어별 기하 요약 — 닫힌 도형 비율과 그 크기."""
    by: dict = {}
    for en in ents:
        t, lay = en.get("t"), str(en.get("l") or "0")
        d = by.setdefault(lay, {"n": 0, "closed": 0, "size": [], "open_len": []})
        d["n"] += 1
        if t == "C":
            d["closed"] += 1
            d["size"].append(float(en.get("r") or 0) * 2)
        elif t == "PL":
            pts = en.get("p") or []
            if len(pts) < 2:
                continue
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            span = math.hypot(max(xs) - min(xs), max(ys) - min(ys))
            shut = math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1])
            if shut <= CLOSE_EPS_MM and len(pts) >= 3:
                d["closed"] += 1
                d["size"].append(span)
            else:
                d["open_len"].append(span)
        elif t == "L":
            p = en.get("p") or []
            if len(p) >= 4:
                d["open_len"].append(math.hypot(p[2] - p[0], p[3] - p[1]))
    return by


def _med(v):
    if not v:
        return 0.0
    v = sorted(v)
    return v[len(v) // 2]


def _report(path):
    from routes.module_f.subdrawing import parse_subdrawing
    ents, _ = parse_subdrawing(path)
    st = _stats(ents)
    print(f"\n[{os.path.basename(path)}]  entity {len(ents):,}")
    print(f"  {'레이어':<22}{'사전':>7}{'수':>7}{'닫힘%':>7}"
          f"{'닫힌크기':>9}{'열린중앙':>9}")
    rows = sorted(st.items(), key=lambda kv: -kv[1]["n"])[:12]
    for lay, d in rows:
        cat = RP._categorize_layer(lay)
        pct = d["closed"] / max(d["n"], 1) * 100
        print(f"  {lay[:22]:<22}{cat:>7}{d['n']:>7}{pct:>6.0f}%"
              f"{_med(d['size']):>9.0f}{_med(d['open_len']):>9.0f}")


if __name__ == "__main__":
    want = sys.argv[1:] or [
        "routes/제출용[최종]/1. 입력도면 대명동 단위세대 계통도.dxf",
        "routes/제출용[최종]/1. 입력도면 대명동 단위세대 기계실.dxf",
        "routes/제출용[최종]/1. 입력도면 대명동 단위세대 평면도.dxf",
        "data/uploads/계통도_LH_306.dxf",
    ]
    for w in want:
        # ★glob 은 «[최종]» 을 문자 클래스로 읽는다 — 이 저장소가 .gitignore
        #   에서 이미 한 번 당한 함정이다. 있는 파일이면 그대로 쓴다.
        hits = [w] if os.path.isfile(w) else glob.glob(glob.escape(w))
        if not hits:
            print(f"\n[{w}] 없음 — 건너뜀")
            continue
        _report(hits[0])
