# -*- coding: utf-8 -*-
"""[F-5] 찍기 후보 제안(D2) + 제외 사유 계측 검증.

수용 기준:
  ① 후보 제안 단계에서 board(찍기 상태)는 불변이다.
  ② 후보 반영은 사람 클릭과 같은 API 로만 들어가고, 반영 수만큼 찍힌 헤드가
     늘며 «찍히지 않음» 집계가 그만큼 준다.
  ③ 인식이 실패해도 찍기는 종전대로 동작한다(부가 기능).

★B1F 로는 못 돌린다 — pick/commit 이 찍은스펙을 공유 작업폴더에 저장하므로
  사용자의 B1F 저장본을 덮어쓴다. 대명동 단위세대 샘플(작업폴더에 그 키가
  없음을 확인)로 돌리고, 만들어진 파일은 끝에서 지운다.

    python tests/test_module_f_suggest.py
"""
from __future__ import annotations

import importlib.util
import io as _io
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
WORK = ROOT / "cad_project_editor_g" / "docs" / "import"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def _wait(c, sid, limit=1200):
    for _ in range(limit):
        jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if jb.get("state") in ("done", "error", "idle"):
            return jb
        time.sleep(0.3)
    return jb


def _cleanup(key: str):
    """이 테스트가 작업폴더에 만든 것만 걷어낸다 — 키 단위로."""
    removed = 0
    for pat in (f"0단계_새찍기/{key}_*", f"0단계_새찍기/{key}*stage1_world*",
                f"DWG/{key}_유저손질.json", f"_edit_disp_cache_{key}.json"):
        for p in WORK.glob(pat):
            try:
                p.unlink()
                removed += 1
            except OSError:
                pass
    print(f"  [정리] 작업폴더에서 {removed}개 걷어냄 (키 {key})")


def _open_and_pick(c, reflect: bool):
    """열기 → 재료 자동 찍기 → 헤드 후보 제안 (+반영) → 상태 반환."""
    with open(DXF, "rb") as f:
        raw = f.read()
    r = c.post("/api/module-f/open", data={
        "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
        content_type="multipart/form-data")
    sid = r.get_json()["sid"]
    assert _wait(c, sid).get("state") == "done", "열기 실패"

    # 재료 — 모듈 A 레이어 사전 자동 찍기(기존 경로)
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
    r = c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
    st = r.get_json()["state"]
    if not st.get("materials"):
        # 자동이 못 잡으면 가장 큰 묶음 중점을 직접 찍는다
        w = c.get(f"/api/module-f/world?sid={sid}").get_json()["world"]
        big = max(w["bundles"], key=lambda b: b["n_seg"])
        sg = big["segs"]
        c.post("/api/module-f/pick/click",
               json={"sid": sid, "x": (sg[0] + sg[2]) / 2,
                     "y": (sg[1] + sg[3]) / 2, "max_d": 5000})
    c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "complete"})
    c.post("/api/module-f/pick/mode",
           json={"sid": sid, "action": "slot", "slot": "상향"})

    # ★world 응답의 찍기 상태 키는 "state" 다 — "pick" 을 읽으면 None 이 와서
    #   불변 비교가 None == None 으로 항상 초록이 된다(실제로 그랬다).
    st0 = c.get(f"/api/module-f/world?sid={sid}").get_json()["state"]
    assert st0, "찍기 상태가 비었다 — 검사가 헛돈다"

    # ── 후보 제안 (board 불변이어야 한다)
    r = c.post("/api/module-f/pick/suggest", json={"sid": sid})
    assert r.get_json().get("ok"), str(r.get_json())
    _wait(c, sid)
    res = c.get(f"/api/module-f/convert/result?sid={sid}").get_json()["result"]
    cands = (res or {}).get("candidates") or []
    st1 = c.get(f"/api/module-f/world?sid={sid}").get_json()["state"]
    return sid, cands, st0, st1


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if not DXF.is_file():
        print(f"샘플 DXF 없음: {DXF}")
        return 1
    key = os.path.splitext(os.path.basename(str(DXF)))[0]
    spec = WORK / "0단계_새찍기" / f"{key}_찍은스펙.json"
    if spec.exists():
        print(f"!! 작업폴더에 이미 {key} 저장본이 있다 — 덮어쓰지 않기 위해 중단")
        return 1

    spec2 = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    srv = importlib.util.module_from_spec(spec2)
    spec2.loader.exec_module(srv)
    app = srv.app
    app.config["TESTING"] = True

    try:
        with app.test_client() as c:
            with c.session_transaction() as s:
                s["authed"] = True

            print("\n[③] 인식 실패/미해당 — 찍기는 종전대로")
            r = c.post("/api/module-f/reopen",
                       json={"key": "B1F 현장조사 소화설비 평면도"})
            rsid = r.get_json()["sid"]
            _wait(c, rsid)
            r = c.post("/api/module-f/pick/suggest", json={"sid": rsid})
            check("DXF 없는 세션(reopen)은 정중히 거절",
                  r.status_code == 400, f"HTTP {r.status_code}")

            print("\n[①] 후보 제안 — board 불변")
            sid, cands, pick_before, pick_after = _open_and_pick(
                c, reflect=False)
            check("후보가 나온다", len(cands) > 0, f"{len(cands)}개")
            check("신뢰도가 실려 있다",
                  all(0 < c_["conf"] <= 1 for c_ in cands),
                  f"최고 {cands[0]['conf'] if cands else '-'}")
            check("제안 전후 찍기 상태 불변(diff 0)",
                  bool(pick_before) and pick_before == pick_after,
                  "제안은 표시일 뿐이다 — 재료 "
                  + str(len((pick_before or {}).get("materials") or [])) + "종")

            print("\n[②] 반영 — 사람 클릭과 같은 API 로만")
            # ★찍기는 «문양 서명» 단위 토글이다 — 한 클릭이 같은 문양 전부를
            #   대표하고, 이미 찍힌 서명 위의 클릭은 «취소» 가 된다. 반영 규칙:
            #   취소로 응답하면 되클릭해 복원하고 «이미 반영» 으로 센다(UI 동일).
            n_before = int((pick_after or {}).get("n_heads") or 0)
            ok_n = dup_n = fail_n = 0
            for c_ in cands:
                d = c.post("/api/module-f/pick/click",
                           json={"sid": sid, "x": c_["x"], "y": c_["y"],
                                 "max_d": 300}).get_json()
                act = (d.get("report") or {}).get("동작")
                if act == "추가":
                    ok_n += 1
                elif act == "취소":
                    c.post("/api/module-f/pick/click",
                           json={"sid": sid, "x": c_["x"], "y": c_["y"],
                                 "max_d": 300})
                    dup_n += 1
                else:
                    fail_n += 1
            st = c.get(f"/api/module-f/world?sid={sid}").get_json()
            n_after = int(((st.get("state") or {}).get("n_heads")) or 0)
            check("새 문양이 반영돼 찍힌 서명이 는다",
                  ok_n > 0 and n_after > n_before,
                  f"서명 {n_before} → {n_after} (새 문양 {ok_n} · "
                  f"이미 찍힘 {dup_n} · E 거름 {fail_n})")
            check("E 의 확정 게이트가 살아 있다(전부 무조건 반영이 아니다)",
                  ok_n + dup_n + fail_n == len(cands),
                  f"후보 {len(cands)} = {ok_n}+{dup_n}+{fail_n}")

            print("\n[②-b] 커밋 → 설계 build 의 제외 사유 집계")
            r = c.post("/api/module-f/pick/commit", json={"sid": sid})
            if not check("커밋 수락", r.get_json().get("ok"),
                         str(r.get_json())[:70]):
                return 1
            if not check("손질망 구성", _wait(c, sid).get("state") == "done"):
                return 1
            # 급수 시작 위치 — 배관 위 아무 점(첫 간선 중점)을 찍는다.
            est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
            seg = est["body_groups"][0]["segs"]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            r = c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                             "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
            est = r.get_json()["state"]
            if not check("급수원 찍힘", len(est.get("sources") or []) == 1,
                         f"{len(est.get('sources') or [])}곳"):
                return 1
            r = c.post("/api/module-f/design/build", json={"sid": sid, "k": 10})
            jb = _wait(c, sid)
            res = c.get(f"/api/module-f/convert/result?sid={sid}"
                        ).get_json()["result"]
            if not check("설계 build", jb.get("state") == "done"
                         and (res or {}).get("ok"),
                         str((res or {}).get("error"))[:90]):
                return 1
            det = res["summary"].get("excluded_detail") or {}
            check("제외 사유 3분류가 집계된다",
                  {"dry", "unpicked"} <= set(det),
                  str(det))
            # «찍히지 않음» 의 정의 검증 — A 후보 중 board 헤드 250mm 안에
            #   없는 것. (A 는 인스턴스를, E 는 문양 서명을 세므로 후보 수와
            #   board 헤드 수는 다른 자다 — E 거름 수와 같아질 이유가 없다.)
            import math as _m
            est2 = c.get(f"/api/module-f/edit/state?sid={sid}"
                         ).get_json()["state"]
            centers = [(h[0], h[1]) for h in est2.get("heads") or []]
            expect_unpicked = sum(
                1 for c_ in cands
                if not any(_m.hypot(c_["x"] - px, c_["y"] - py) <= 250.0
                           for px, py in centers))
            check("찍히지 않음 == A 후보 중 board 에 없는 것",
                  det.get("unpicked") == expect_unpicked,
                  f"집계 {det.get('unpicked')} vs 재계산 {expect_unpicked} "
                  f"(후보 {len(cands)} · board 헤드 {len(centers)})")
            pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
            marks = pv.get("marks") or {}
            check("분류 좌표가 화면용으로 실린다",
                  all(isinstance(marks.get(k2, {}).get("xy"), list)
                      for k2 in det), str(list(marks.keys())))
    finally:
        _cleanup(key)

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("F-5 찍기 후보 제안 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
