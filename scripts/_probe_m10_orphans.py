# -*- coding: utf-8 -*-
"""[M10 · M8] 표가 가리키는 배관 라벨이 배관표에 **전부** 있는가 — 고아 0.

§3-6 이 배관 라벨을 다시 매기면서, 라벨을 쓰는 자리가 여럿이 됐다:
    배관표 · 부속표 · 기기표(`pipe` 칸) · 미리보기 응답 · .sdf/.kfp/.has
한 자리라도 **옛 이름**을 들고 있으면 그 행은 배관표에서 못 찾는 «고아» 가 된다.
조용히 빈칸으로 새지 그 자리에서 터지지 않는다 — 그래서 세어야 한다.

M8 도 함께 잰다: 새 기기의 등가길이가 라이브러리에 없으면 0 이 아니라
«미해결» 로 세어지는가.

    python scripts/_probe_m10_orphans.py [--key ...] [--k 30]
"""
from __future__ import annotations

import argparse
import contextlib
import importlib
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
for p in (ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)

DM_KEY = "1. 입력도면 대명동 단위세대 평면도"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def run(key, k) -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from _probe_candidate_drop import setup, wait
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    buf = io.StringIO()
    print(f"■ M10 고아 · M8 미해결 — {key} · K={k}")
    with isolated_workdir(prefix="m10_", copy_key=key), \
            srv.app.test_client() as c:
        with contextlib.redirect_stdout(buf):
            with c.session_transaction() as s:
                s["authed"] = True
            sid = setup(c, argparse.Namespace(key=key))
            st = c.get(f"/api/module-f/edit/state?sid={sid}"
                       ).get_json()["state"]
            if not (st.get("sources") or ()):
                seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                             key=len, reverse=True)[0]
                c.post("/api/module-f/edit/mode",
                       json={"sid": sid, "mode": "급수시작위치"})
                c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                             "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
            c.post("/api/module-f/edit/worst", json={"sid": sid, "k": k})
            c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
            if wait(c, sid).get("state") != "done":
                print("★표 확정 실패")
                return 1
            pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()

        tabs = pv.get("tables") or {}
        pipes = tabs.get("pipes") or []
        labels = {str(r.get("label")) for r in pipes}
        print(f"  표: 배관 {len(pipes)} · 절점 {len(tabs.get('nodes') or [])} "
              f"· 부속 {len(tabs.get('fittings') or [])} "
              f"· 기기 {len(tabs.get('equipment') or [])}")

        # ── M10 ① 부속표·기기표가 가리키는 배관이 배관표에 있는가
        for name, rows, field in (("부속표", tabs.get("fittings") or [], "pipe"),
                                  ("기기표", tabs.get("equipment") or [], "pipe")):
            refs = [str(r.get(field)) for r in rows if r.get(field) is not None]
            orphan = sorted({x for x in refs if x not in labels})
            check(f"M10 {name} 고아 0", not orphan,
                  f"참조 {len(refs)} · 고아 {len(orphan)}"
                  + (f" 예: {orphan[:5]}" if orphan else ""))

        # ── M10 ② 미리보기 응답의 배관 라벨도 같은 이름공간인가
        vpipes = ((pv.get("view") or {}).get("pipes")) or []
        vlab = {str(p.get("label")) for p in vpipes}
        check("M10 미리보기 라벨 == 배관표 라벨", vlab == labels,
              f"미리보기 {len(vlab)} · 표 {len(labels)} · "
              f"표에 없는 것 {sorted(vlab - labels)[:5]}")

        # ── M10 ③ 라벨이 중복되지 않는가(재번호가 겹치면 조용히 덮는다)
        check("M10 배관 라벨 중복 0", len(labels) == len(pipes),
              f"고유 {len(labels)} / 행 {len(pipes)}")

        # ── M10 ④ 절점표와 배관 양끝(in/out)이 맞물리는가
        nlab = {str(r.get("label")) for r in (tabs.get("nodes") or [])}
        bad_ends = sorted({str(r.get(e)) for r in pipes for e in ("in", "out")
                           if str(r.get(e)) not in nlab})
        check("M10 배관 양끝이 절점표에 있다", not bad_ends,
              f"고아 끝 {len(bad_ends)}" + (f" 예: {bad_ends[:5]}"
                                          if bad_ends else ""))

        # ── M8 등가길이 미해결이 0 으로 새지 않는가
        meta = dict(tabs.get("meta") or [])
        unres = meta.get("등가길이 미해결")
        eqs = [r.get("eq_len") for r in (tabs.get("equipment") or [])]
        zero_eq = sum(1 for v in eqs if v == 0)
        check("M8 미해결이 집계로 뜬다", unres is not None,
              f"등가길이 미해결 = {unres}")
        print(f"       기기 등가길이 값: {eqs} · 0 으로 적힌 것 {zero_eq}개")

    print()
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"실패 {len(FAILS)}건")
        return 3
    print("★M10 고아 0 · M8 미해결 집계 — 선다")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=DM_KEY)
    ap.add_argument("--k", type=int, default=30)
    a = ap.parse_args()
    return run(a.key, a.k)


if __name__ == "__main__":
    raise SystemExit(main())
