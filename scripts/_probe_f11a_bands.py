# -*- coding: utf-8 -*-
"""[F-11a] 지배 띠 채택이 세 도면에서 실제로 무엇을 바꾸는가.

수용 기준이 도면별로 다르다:

    대명동  높음만 채택 — 종전과 «같은 집합» (회귀 없음)
    B1F     높음+중간 채택 → 원클릭 뒤 최불리 30개 (2개 퇴화 해소)
    LH306   높음+중간 40개 채택 → 조립이 산다 (§16 게이트 미발동)

화면 규칙(JS)이 아니라 **서버가 내려주는 임계** 로 흐름을 태워 잰다 — 규칙이
한 곳에만 있는지도 이걸로 같이 확인된다.

    python scripts/_probe_f11a_bands.py [도면.dxf ...]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
]


def wait(c, sid, limit=9000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def run(c, dxf):
    with open(dxf, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {})["sid"]
    wait(c, sid)
    rec = (c.get(f"/api/module-f/recon?sid={sid}").get_json()
           or {}).get("recon") or {}
    ad = rec.get("adopt") or {}
    print(f"\n■ {dxf.name}")
    print(f"    띠 {rec.get('bands')}")
    print(f"    규칙 {ad.get('rule')} · 임계 {ad.get('conf_min')} · "
          f"채택 예정 {ad.get('n')}")
    print(f"    → {ad.get('why')}")
    if ad.get("conf_min") is None:
        print("    ★규칙이 0 을 냈다 — 게이트가 찍기로 보낸다")
        return None

    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True,
                 "heads": {"conf_min": ad["conf_min"]}})
    j = wait(c, sid)
    res = (c.get(f"/api/module-f/convert/result?sid={sid}")
           .get_json() or {}).get("result") or {}
    print(f"    채택 {j['state']} · 헤드 찍힘 {res.get('head_applied')} · "
          f"이미 반영 {res.get('head_already')} · 유령 {res.get('head_skipped')}")
    r = c.post("/api/module-f/pick/commit", json={"sid": sid})
    j = wait(c, sid)
    print(f"    조립 {j['state']}" + (f" — {j.get('error')}" if j.get("error") else ""))
    if j["state"] != "done":
        return sid
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    print(f"    board 노드 {st['counts']['pts']:,} · 간선 "
          f"{st['counts']['edges']:,} · 헤드 {st['counts']['heads']:,}")

    # ── 원클릭. ★앵커 자리를 아무 데나 고르면 안 된다 — B1F 는 배관이 306
    #    조각이라 작은 조각에 걸리면 최불리가 «1개 · 0.18m» 로 나온다. 그건
    #    띠 규칙의 문제가 아니라 «어디를 찍었나» 의 문제다. 사람도 주배관을
    #    보고 찍으므로 큰 덩이부터 후보를 만들어 가장 좋은 것을 쓴다
    #    (F-10g 완주 스크립트와 같은 방식).
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    groups = sorted((g.get("segs") or [] for g in st["body_groups"]),
                    key=len, reverse=True)
    cands = []
    for s2 in groups[:18]:
        pts = []
        for i in range(0, len(s2) - 3, 4):
            pts.append((float(s2[i]), float(s2[i + 1])))
        if not (pts and heads):
            continue
        best, bd = None, None
        for hx, hy in heads:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if bd is None or d < bd:
                best, bd = p, d
        if best is not None and bd <= 2000.0:
            cands.append(best)
        if len(cands) >= 6:
            break
    if not cands:
        print("    (찍을 자리를 못 찾아 원클릭 생략)")
        return sid
    top, tries = None, 0
    for (px, py) in cands:
        tries += 1
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        j = wait(c, sid)
        st2 = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        w = st2.get("worst")
        if w and (top is None or int(w["k"]) > int(top["k"])):
            top = w
        if top and int(top["k"]) >= 30:
            break
    print(f"    원클릭 · 찍은 횟수 {tries} · 최불리 "
          + (f"{top['k']}개 · 최원 {top['far_m']} m" if top else "없음"))
    return sid


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        for f in [Path(x) for x in sys.argv[1:]] or DEFAULT:
            if not f.is_file():
                print(f"\n■ {f.name} — 파일 없음")
                continue
            try:
                run(c, f)
            except Exception as exc:  # noqa: BLE001
                print(f"    ★실패 {type(exc).__name__}: {exc}")
    print("\n  대명동=높음만 · B1F/LH306=중간까지 여야 한다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
