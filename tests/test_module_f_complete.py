# -*- coding: utf-8 -*-
"""[F-7] 완성품 F 골든 회귀 — 열기 → 찍기 → 손질 → 최불리 → 3종 산출 전 구간.

두 구간으로 나눈다:
  Ⅰ. B1F(저장본 reopen) — F-1 기준선과 수치 대조 + 3종 산출 + SDF 불변식.
     찍기 커밋은 하지 않는다(사용자 저장본을 덮는다).
  Ⅱ. 대명동 샘플(전 구간) — 열기부터 커밋·산출까지 실제로 태우고 골든과 대조.
     작업폴더에 만든 것은 끝에서 걷어낸다.

기준선 원칙(지시서 §0.4): F-1 이전 산출물은 쓰지 않는다. 골든에는 board 지문을
함께 record 해 «입력 변화» 와 «코드 회귀» 를 갈라 말한다(G B6 원칙).

    python tests/test_module_f_complete.py
    python tests/test_module_f_complete.py --record   # 골든 갱신(의도 시에만)
"""
from __future__ import annotations

import importlib.util
import io as _io
import json
import os
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
B1F = "B1F 현장조사 소화설비 평면도"
DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
WORK = ROOT / "cad_project_editor_g" / "docs" / "import"
GOLDEN = ROOT / "tests" / "module_f_complete_golden.json"
BASELINE = ROOT / "tests" / "module_f_worst_baseline.json"
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


def _sdf_invariants(path: Path) -> dict:
    """SDF 불변식 — G14~G19 가 지키는 것들을 파일에서 직접 센다."""
    txt = path.read_text(encoding="utf-8")
    head = txt.splitlines()[:2]
    r = ET.fromstring(txt)
    sets = [ps.findtext("Pipe-type/Name") for ps in r.iter("Pipe-set")]
    libs = [ul.get("file") for lb in r.iter("Libraries")
            for ul in lb.findall("User-lib")]
    no_type = unbound = 0
    for links in r.iter("Links"):
        for ps in links.findall("Pipe-set"):
            nm = ps.findtext("Pipe-type/Name")
            sizes = {round(float(e.get("size")), 6)
                     for e in ps.findall("Pipe-type/Pipe-size")}
            for pipe in ps.findall("Pipe"):
                if not nm:
                    no_type += 1
                elif round(float(pipe.get("bore") or 0), 6) not in sizes:
                    unbound += 1
    return {
        "doctype": 'DOCTYPE Project SYSTEM "spray.dtd"' in (head[1]
                                                            if len(head) > 1
                                                            else ""),
        "pipe_set_names": [n for n in sets if n],
        "placeholder_first": sets[0] is None if sets else False,
        "user_lib": libs,
        "user_lib_relative": all("\\" not in (u or "") and "/" not in (u or "")
                                 for u in libs),
        "none_defined": no_type, "unset_bore": unbound,
        "nodes": len(list(r.iter("Node"))),
        "pipes": len(list(r.iter("Pipe"))),
        "nozzles": len(list(r.iter("Nozzle"))),
    }


def _cleanup(key: str):
    removed = 0
    for pat in (f"0단계_새찍기/{key}_*", f"0단계_새찍기/{key}*stage1_world*",
                f"DWG/{key}_유저손질.json", f"_edit_disp_cache_{key}.json"):
        for p in WORK.glob(pat):
            try:
                p.unlink()
                removed += 1
            except OSError:
                pass
    print(f"  [정리] 작업폴더 {removed}개 걷어냄 (키 {key})")


def part1_b1f(c, record: dict) -> None:
    print("\n[Ⅰ] B1F 저장본 — F-1 기준선 대조 + 3종 산출")
    sid = c.post("/api/module-f/reopen", json={"key": B1F}).get_json()["sid"]
    if not check("reopen", _wait(c, sid).get("state") == "done"):
        return
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    board = {"pts": st["counts"]["pts"], "edges": st["counts"]["edges"],
             "heads": st["counts"]["heads"]}

    j = c.post("/api/module-f/edit/worst",
               json={"sid": sid, "k": 30}).get_json()
    s = j["summary"]
    base = json.loads(BASELINE.read_text(encoding="utf-8"))
    if base.get("board") != board:
        print(f"  [정보] board 가 기준선과 다르다 — {base.get('board')} → "
              f"{board}. 입력 변화다(코드 회귀 아님). --record 로 다시 뜨라.")
        FAILS.append("B1F board 지문 불일치(입력 변화)")
        return
    for k2 in ("far_m", "near_m", "span_m", "total_m", "max_load"):
        check(f"F-1 기준선 {k2}", s.get(k2) == base.get(k2),
              f"기준 {base.get(k2)} / 지금 {s.get(k2)}")
    check("급수원 기준 표기", s.get("source") == base.get("source"),
          str(s.get("source")))

    # 설계 표 + 3종 산출
    c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
    if not check("design build", _wait(c, sid).get("state") == "done"):
        return
    c.post("/api/module-f/convert/run",
           json={"sid": sid, "dto": {}, "outputs": {
               "full_kfp": True, "worst_kfp": True, "worst_sdf": True}})
    if not check("3종 산출", _wait(c, sid).get("state") == "done"):
        return
    res = c.get(f"/api/module-f/convert/result?sid={sid}").get_json()["result"]
    sm = res["summary"]
    cur = {"board": board,
           "worst": {k2: s.get(k2) for k2 in
                     ("far_m", "span_m", "total_m", "max_load")},
           "full": {"nodes": sm["full"]["nodes"], "pipes": sm["full"]["pipes"]},
           "worst_kfp": {"k": sm["worst"]["k"], "nodes": sm["worst"]["nodes"],
                         "pipes": sm["worst"]["pipes"]}}
    sdf = next((ROOT / "data" / "uploads" / "module_f").glob(
        f"{sid}_design/*.sdf"))
    inv = _sdf_invariants(sdf)
    cur["sdf"] = {k2: inv[k2] for k2 in
                  ("pipe_set_names", "placeholder_first", "nodes", "pipes",
                   "nozzles")}
    check("SDF 불변식 — DOCTYPE·placeholder·상대 SLF 참조",
          inv["doctype"] and inv["placeholder_first"]
          and inv["user_lib_relative"] and len(inv["user_lib"]) == 1,
          f"libs {inv['user_lib']}")
    check("SDF 불변식 — Type/Diameter 빈 곳 0",
          inv["none_defined"] == 0 and inv["unset_bore"] == 0,
          f"none {inv['none_defined']} · unset {inv['unset_bore']}")
    record["b1f"] = cur


def part2_full_path(c, record: dict) -> None:
    print("\n[Ⅱ] 대명동 — 열기 → 찍기 → 손질 → 최불리 → 산출 전 구간")
    key = os.path.splitext(os.path.basename(str(DXF)))[0]
    spec = WORK / "0단계_새찍기" / f"{key}_찍은스펙.json"
    if spec.exists():
        FAILS.append(f"작업폴더에 이미 {key} 저장본 — 전 구간 골든 생략")
        return
    try:
        with open(DXF, "rb") as f:
            raw = f.read()
        r = c.post("/api/module-f/open", data={
            "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
            content_type="multipart/form-data")
        sid = r.get_json()["sid"]
        check("열기", _wait(c, sid).get("state") == "done")
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid, "action": "complete"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid, "action": "slot", "slot": "상향"})
        # 헤드 — 후보 제안을 반영(취소면 되클릭 복원, F-5 규칙 그대로)
        c.post("/api/module-f/pick/suggest", json={"sid": sid})
        _wait(c, sid)
        cands = (c.get(f"/api/module-f/convert/result?sid={sid}"
                       ).get_json()["result"] or {}).get("candidates") or []
        for c_ in cands:
            d = c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300}).get_json()
            if (d.get("report") or {}).get("동작") == "취소":
                c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300})
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        check("커밋 → 손질망", _wait(c, sid).get("state") == "done")
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        seg = est["body_groups"][0]["segs"]
        c.post("/api/module-f/edit/mode",
               json={"sid": sid, "mode": "급수시작위치"})
        c.post("/api/module-f/edit/click",
               json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                     "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
        j = c.post("/api/module-f/edit/worst",
                   json={"sid": sid, "k": 10}).get_json()
        s = j["summary"]
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 10})
        check("design build", _wait(c, sid).get("state") == "done")
        c.post("/api/module-f/convert/run",
               json={"sid": sid, "dto": {}, "outputs": {
                   "full_kfp": True, "worst_kfp": True, "worst_sdf": True}})
        check("3종 산출", _wait(c, sid).get("state") == "done")
        res = c.get(f"/api/module-f/convert/result?sid={sid}"
                    ).get_json()["result"]
        sm = res["summary"]
        est2 = c.get(f"/api/module-f/edit/state?sid={sid}"
                     ).get_json()["state"]
        record["daemyeong"] = {
            "board": {"pts": est2["counts"]["pts"],
                      "edges": est2["counts"]["edges"],
                      "heads": est2["counts"]["heads"]},
            "worst": {k2: s.get(k2) for k2 in
                      ("far_m", "span_m", "total_m", "max_load")},
            "full": {"nodes": sm["full"]["nodes"],
                     "pipes": sm["full"]["pipes"]},
            "worst_kfp": {"k": sm["worst"]["k"],
                          "nodes": sm["worst"]["nodes"],
                          "pipes": sm["worst"]["pipes"]},
            "design": {"sdf": sm["design"]["sdf"]},
        }
    finally:
        _cleanup(key)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    spec2 = importlib.util.spec_from_file_location(
        "daejo", os.path.join(str(ROOT), "대조 서버.py"))
    srv = importlib.util.module_from_spec(spec2)
    spec2.loader.exec_module(srv)
    app = srv.app
    app.config["TESTING"] = True

    record: dict = {}
    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        part1_b1f(c, record)
        part2_full_path(c, record)

    print("\n[골든 대조]")
    if "--record" in sys.argv or not GOLDEN.is_file():
        GOLDEN.write_text(json.dumps(record, ensure_ascii=False, indent=2),
                          encoding="utf-8")
        print(f"  [기록] {GOLDEN.name} 갱신")
    else:
        gold = json.loads(GOLDEN.read_text(encoding="utf-8"))
        for part in ("b1f", "daemyeong"):
            if part not in gold or part not in record:
                continue
            g, cur = gold[part], record[part]
            if g.get("board") != cur.get("board"):
                print(f"  [정보] {part} board 표류 — {g.get('board')} → "
                      f"{cur.get('board')} (입력 변화 · --record 로 재기록)")
                FAILS.append(f"{part} board 지문 불일치(입력 변화)")
                continue
            check(f"골든 {part}", g == cur,
                  json.dumps({k2: (g.get(k2), cur.get(k2))
                              for k2 in g if g.get(k2) != cur.get(k2)},
                             ensure_ascii=False)[:160] or "동일")

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("F-7 완성품 골든 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
