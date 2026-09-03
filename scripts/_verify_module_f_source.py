# -*- coding: utf-8 -*-
"""[F-1] 급수원 지정 규약 검증 — D4 · G BLOCKED B2 해소.

수용 기준 네 가지를 실측한다:
  ① 1곳이면 자동(Z1) — 미지정 호출이 그대로 통한다.
  ② 2곳 이상 + 미지정 → 400 과 후보 목록(kfp 변환의 source_selection_required 규약).
  ③ Z1 지정과 Z2 지정의 앵커·최원 유하거리가 서로 다르고, 각각 재실행 시 재현된다.
  ④ 그 결과가 F-1 기준선(tests/module_f_worst_baseline.json)과 같다.

두 번째 급수원은 저장본을 건드리지 않고 **세션 안에서만** 찍는다(찍고 지우는
것도 같은 API 다 — 저장하지 않으면 파일은 불변).

    python scripts/_verify_module_f_source.py
    python scripts/_verify_module_f_source.py --record   # 기준선 갱신(의도 시에만)
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEY = "B1F 현장조사 소화설비 평면도"
BASELINE = ROOT / "tests" / "module_f_worst_baseline.json"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    spec = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    srv = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(srv)
    app = srv.app
    app.config["TESTING"] = True

    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True

        r = c.post("/api/module-f/reopen", json={"key": KEY})
        sid = r.get_json()["sid"]
        for _ in range(300):
            jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
            if jb.get("state") in ("done", "error"):
                break
            time.sleep(0.3)
        if not check("B1F reopen", jb.get("state") == "done", str(jb)[:80]):
            return 1
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        n_src = len(st.get("sources") or [])
        check("급수원 1곳에서 출발", n_src == 1, f"{n_src}곳")

        def worst(source=None):
            body = {"sid": sid, "k": 30}
            if source is not None:
                body["source"] = source
            return c.post("/api/module-f/edit/worst", json=body)

        print("\n[①] 1곳이면 자동")
        j = worst().get_json()
        check("미지정 호출이 통한다(자동 Z1)",
              j.get("ok") and j["summary"].get("source") == "Z1",
              f"source={j.get('summary', {}).get('source')}")
        z1_alone = dict(j["summary"])

        print("\n[②] 두 번째 급수원을 세션에만 찍는다")
        # 앵커 반대쪽 임의 배관 위 — 저장하지 않으므로 파일은 불변이다.
        before_mtime = None
        edits = (ROOT / "cad_project_editor_g" / "docs" / "import" / "DWG"
                 / f"{KEY}_유저손질.json")
        if edits.is_file():
            before_mtime = edits.stat().st_mtime
        c.post("/api/module-f/edit/mode",
               json={"sid": sid, "mode": "급수시작위치"})
        anchor = (st.get("worst") or {}).get("worst_head")
        # 화면에 있는 아무 헤드 하나를 급수원 후보 좌표로 쓴다 — 배관에 스냅된다.
        heads = st.get("heads") or []
        far_head = heads[len(heads) // 2]
        r = c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": far_head[0], "y": far_head[1],
                         "max_d": 500.0})
        st2 = r.get_json()["state"]
        n_src2 = len(st2.get("sources") or [])
        check("급수원이 2곳이 됐다(세션 안)", n_src2 == 2, f"{n_src2}곳")

        j = worst()
        check("미지정 → 400", j.status_code == 400, f"HTTP {j.status_code}")
        jj = j.get_json()
        check("규약 코드 — source_selection_required",
              jj.get("code") == "source_selection_required", str(jj.get("code")))
        check("후보 목록(Z1·Z2)이 온다",
              [c_.get("tag") for c_ in (jj.get("sources") or [])] == ["Z1", "Z2"],
              str(jj.get("sources"))[:80])
        j = worst("Z9")
        check("없는 태그도 400 + 후보", j.status_code == 400
              and len(j.get_json().get("sources") or []) == 2,
              f"HTTP {j.status_code}")

        print("\n[③] Z1 과 Z2 는 다른 답이고, 각각 재현된다")
        a1 = worst("Z1").get_json()["summary"]
        a2 = worst("Z2").get_json()["summary"]
        check("Z1 결과에 기준 급수원이 적혀 있다", a1.get("source") == "Z1",
              str(a1.get("source")))
        check("Z1 ≠ Z2 (최원 유하거리)", a1["far_m"] != a2["far_m"],
              f"Z1 {a1['far_m']} m vs Z2 {a2['far_m']} m")
        b1 = worst("Z1").get_json()["summary"]
        b2 = worst("Z2").get_json()["summary"]
        check("Z1 재실행 재현", a1 == b1, f"{a1['far_m']} m")
        check("Z2 재실행 재현", a2 == b2, f"{a2['far_m']} m")
        check("1곳 시절 자동 결과 == Z1 지정 결과(급수원 추가 전 기준)",
              z1_alone["far_m"] == a1["far_m"]
              and z1_alone["max_load"] == a1["max_load"],
              f"{z1_alone['far_m']} m vs {a1['far_m']} m")

        if before_mtime is not None:
            check("저장본 파일은 불변(세션에만 찍었다)",
                  edits.stat().st_mtime == before_mtime, str(edits.name))

        print("\n[④] F-1 기준선")
        cur = {"key": KEY, "source": "Z1",
               "board": {"pts": st["counts"]["pts"],
                         "edges": st["counts"]["edges"],
                         "heads": st["counts"]["heads"]},
               "k": a1["k"], "far_m": a1["far_m"], "near_m": a1["near_m"],
               "span_m": a1["span_m"], "total_m": a1["total_m"],
               "max_load": a1["max_load"]}
        if "--record" in sys.argv or not BASELINE.is_file():
            BASELINE.write_text(
                json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")
            print(f"  [기록] {BASELINE.name} ← {cur}")
        else:
            old = json.loads(BASELINE.read_text(encoding="utf-8"))
            same = all(old.get(k) == cur.get(k) for k in
                       ("far_m", "near_m", "span_m", "total_m", "max_load", "k"))
            if not same and old.get("board") != cur.get("board"):
                # 입력이 달라진 것과 코드 회귀를 갈라 말한다(B6 과 같은 원칙).
                print(f"  [정보] board 가 달라졌다 — {old.get('board')} → "
                      f"{cur.get('board')}. 코드 회귀가 아니라 입력 변화다. "
                      f"--record 로 기준선을 다시 뜨라.")
                FAILS.append("기준선 board 불일치(입력 변화)")
            else:
                check("F-1 기준선과 일치", same,
                      f"기준 {old.get('far_m')} m / 지금 {cur.get('far_m')} m")

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("F-1 급수원 지정 규약 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
