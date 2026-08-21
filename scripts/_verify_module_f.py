# -*- coding: utf-8 -*-
"""모듈 F 전 경로 런타임 검증 — 테스트 클라이언트로 4단을 실제로 태운다.

라우트가 등록됐는지만 보는 인벤토리로는 부족하다. 업로드 → 찍기 클릭 →
배관망 구성 → 손질 → 변환 → 내려받기까지 실제 응답 본문을 확인한다.
"""
from __future__ import annotations

import importlib.util
import io
import json
import os
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

spec = importlib.util.spec_from_file_location("daejo", os.path.join(ROOT, "대조 서버.py"))
srv = importlib.util.module_from_spec(spec)
sys.modules["daejo"] = srv
spec.loader.exec_module(srv)
app = srv.app
app.config["TESTING"] = True

DXF = r"C:\Users\admin\Desktop\B1F 현장조사 소화설비 평면도.dxf"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def wait(c, sid, what, limit=600):
    t0 = time.time()
    last = ""
    while time.time() - t0 < limit:
        r = c.get(f"/api/module-f/job?sid={sid}")
        j = r.get_json()
        if j.get("state") in ("done", "error"):
            print(f"      {what} {j['state']} · {j['elapsed']}s")
            if j["state"] == "error":
                print("      ", j.get("error"))
                for ln in j.get("lines", [])[-6:]:
                    print("       |", ln)
            return j
        cur = (j.get("lines") or [""])[-1][:78]
        if cur != last:
            print(f"      … {cur}")
            last = cur
        time.sleep(0.6)
    return {"state": "timeout"}


def main():
    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True

        print("\n[0] 페이지 · 저장목록 · 변환 폼")
        r = c.get("/module-f")
        check("GET /module-f", r.status_code == 200 and b"MODULE F" in r.data,
              f"HTTP {r.status_code} {len(r.data)}B")
        r = c.get("/api/module-f/saved")
        j = r.get_json()
        check("saved 목록", j.get("ok") and isinstance(j.get("items"), list),
              f"{len(j.get('items') or [])}건")
        r = c.get("/api/module-f/convert/fields")
        j = r.get_json()
        nf = sum(len(g["fields"]) for g in (j.get("groups") or []))
        check("변환 폼 필드", j.get("ok") and nf >= 12, f"{nf}칸")

        print("\n[1] 만료 세션 방어")
        r = c.get("/api/module-f/world?sid=nonexistent")
        check("없는 sid → 410", r.status_code == 410, f"HTTP {r.status_code}")

        if not os.path.isfile(DXF):
            print(f"\n!! 시험용 DXF 없음: {DXF} — 2~4단 생략")
            return

        print("\n[2] DXF 열기 → 찍기")
        with open(DXF, "rb") as f:
            raw = f.read()
        r = c.post("/api/module-f/open", data={
            "dxf_file": (io.BytesIO(raw), os.path.basename(DXF))},
            content_type="multipart/form-data")
        j = r.get_json()
        if not check("업로드 수락", j.get("ok"), str(j)[:200]):
            return
        sid = j["sid"]
        if wait(c, sid, "도면 읽기")["state"] != "done":
            FAILS.append("도면 읽기 실패")
            return

        w = c.get(f"/api/module-f/world?sid={sid}").get_json()
        wd = w["world"]
        check("세계 도형", wd["counts"]["segs"] > 1000,
              f"선분 {wd['counts']['segs']} · 묶음 {len(wd['bundles'])}"
              f" · 표시생략 {wd['dropped']}")
        check("경계 유효", wd["bounds"]["maxx"] > wd["bounds"]["minx"],
              json.dumps(wd["bounds"])[:90])
        payload_bytes = len(json.dumps(wd))
        check("world JSON 크기", payload_bytes < 40_000_000,
              f"{payload_bytes/1024/1024:.1f} MB")

        # 배관 재료 찍기 — 실제 선분 위 중점을 노려 찍는다.
        big = max(wd["bundles"], key=lambda b: b["n_seg"])
        sg = big["segs"]
        mx = (sg[0] + sg[2]) / 2
        my = (sg[1] + sg[3]) / 2
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        r = c.post("/api/module-f/pick/click",
                   json={"sid": sid, "x": mx, "y": my, "max_d": 5000})
        j = r.get_json()
        check("배관 클릭", bool(j.get("report")), str(j.get("report"))[:110])
        check("강조 반영", len(j["state"]["highlight"]["pipe_segs"]) > 0,
              f"{len(j['state']['highlight']['pipe_segs'])}선")
        r = c.post("/api/module-f/pick/undo", json={"sid": sid})
        check("되돌리기", r.get_json()["state"]["materials"] == [],
              "재료 0개로 복귀")
        c.post("/api/module-f/pick/click",
               json={"sid": sid, "x": mx, "y": my, "max_d": 5000})
        r = c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "complete"})
        j = r.get_json()
        check("선택완료 → 헤드 모드", j["state"]["mat_done"] and j["state"]["mode"] == "헤드",
              j["message"])
        r = c.post("/api/module-f/pick/mode",
                   json={"sid": sid, "action": "slot", "slot": "상하향"})
        check("헤드 칸 전환", r.get_json()["state"]["head_label"] == "상하향",
              r.get_json()["state"]["head_label"])

        print("\n[3] 저장본으로 손질 (찍기 재구성은 별도 세션에서)")
        r = c.post("/api/module-f/reopen",
                   json={"key": "B1F 현장조사 소화설비 평면도"})
        j = r.get_json()
        if not check("reopen 수락", j.get("ok"), str(j)[:160]):
            return
        sid2 = j["sid"]
        if wait(c, sid2, "배관망 열기")["state"] != "done":
            FAILS.append("배관망 열기 실패")
            return
        st = c.get(f"/api/module-f/edit/state?sid={sid2}").get_json()["state"]
        check("손질 망", st["counts"]["heads"] > 0,
              f"노드 {st['counts']['pts']} · 간선 {st['counts']['edges']}"
              f" · 헤드 {st['counts']['heads']} · 덩이 {st['counts']['bodies']}")
        check("헤드 종류 집계", bool(st["kinds"]), json.dumps(st["kinds"], ensure_ascii=False))
        check("표시 좌표", len(st["body_groups"]) > 0 and len(st["heads"]) > 0,
              f"덩이 {len(st['body_groups'])} · 헤드점 {len(st['heads'])}")

        for mode in ("삭제", "급수시작위치", "알람밸브위치", "이음"):
            r = c.post("/api/module-f/edit/mode", json={"sid": sid2, "mode": mode})
            if not check(f"모드 {mode}", r.get_json()["state"]["mode"] == mode):
                break
        r = c.post("/api/module-f/edit/mode", json={"sid": sid2, "mode": "없는모드"})
        check("모르는 모드 거절", r.status_code == 400, f"HTTP {r.status_code}")

        r = c.post("/api/module-f/edit/flow", json={"sid": sid2})
        j = r.get_json()
        if j.get("ok"):
            check("물흐름", j["water"]["wet_heads"] > 0,
                  f"{j['water']['wet_heads']}/{j['water']['total_heads']} 헤드"
                  f" · 젖은간선 {j['water']['wet_edges']}")
            check("물길 좌표 유지", len(j["state"]["wet_pipes"]) > 0,
                  f"{len(j['state']['wet_pipes'])}개")
            again = c.get(f"/api/module-f/edit/state?sid={sid2}").get_json()["state"]
            check("물길이 다시 조회해도 남음", len(again["wet_pipes"]) > 0,
                  f"{len(again['wet_pipes'])}개")
        else:
            check("물흐름", False, j.get("message"))

        r = c.post("/api/module-f/edit/kind", json={"sid": sid2, "kind": "상향식"})
        check("헤드 미선택 시 종류변경 거절", r.status_code == 400,
              f"HTTP {r.status_code}")

        print("\n[3-A] 모듈 A 이식 — 레이어 자동 추천")
        cats = wd.get("cats") or {}
        check("레이어 카테고리 분류", bool(cats),
              json.dumps(cats, ensure_ascii=False))
        check("배관 추천 존재", cats.get("PIPE", 0) > 0, f"PIPE {cats.get('PIPE')}묶음")
        check("묶음마다 cat 부여",
              all("cat" in b for b in wd["bundles"]), f"{len(wd['bundles'])}묶음")
        r = c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        j = r.get_json()
        check("PIPE 추천 일괄 찍기", j.get("ok") and len(j["applied"]) > 0,
              f"{len(j.get('applied') or [])}묶음 · {j.get('message')}")
        check("찍기판에 반영", j["state"]["materials"],
              f"재료 {len(j['state']['materials'])}묶음")
        r = c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "TEXT"})
        check("추천 카테고리 아닌 값 거절", r.status_code == 400,
              f"HTTP {r.status_code}")

        print("\n[3-B] 모듈 A 이식 — Remote 30 최불리 헤드")
        r = c.post("/api/module-f/edit/worst", json={"sid": sid2, "k": 30})
        j = r.get_json()
        if check("최불리 선정", j.get("ok"), str(j.get("message"))[:120]):
            s = j["summary"]
            check("30개 선정", s["k"] == 30,
                  f"{s['k']}개 / 도달 {s['reachable']}")
            check("거리 순서(최원 ≥ 끝)", s["far_m"] >= s["near_m"],
                  f"최원 {s['far_m']} m · 끝 {s['near_m']} m")
            check("경로 간선 축소", 0 < s["path_edges"] < st["counts"]["edges"],
                  f"{s['path_edges']} / 전체 {st['counts']['edges']}")
            w = j["state"]["worst"]
            check("화면 좌표 동봉", len(w["heads"]) == 30 and len(w["path"]) > 0,
                  f"헤드 {len(w['heads'])} · 경로 {len(w['path'])}")
        r = c.post("/api/module-f/edit/worst-clear", json={"sid": sid2})
        check("선정 해제", r.get_json()["state"]["worst"] is None)
        c.post("/api/module-f/edit/worst", json={"sid": sid2, "k": 30})

        print("\n[4] 변환 → 내려받기")
        r = c.post("/api/module-f/convert/run", json={"sid": sid2, "dto": {}})
        if not check("변환 요청 수락", r.get_json().get("ok"), str(r.get_json())[:160]):
            return
        if wait(c, sid2, "KFP 변환")["state"] != "done":
            FAILS.append("변환 잡 실패")
            return
        res = c.get(f"/api/module-f/convert/result?sid={sid2}").get_json()["result"]
        if not check("변환 성공", res.get("ok"),
                     json.dumps(res.get("blockers"), ensure_ascii=False)[:220]):
            return
        s = res["summary"]
        check("KFP 내용", s["nodes"] > 0 and s["pipes"] > 0,
              f"노드 {s['nodes']} · 배관 {s['pipes']} · {s['bytes']:,}B")
        sdf = s.get("sdf") or {}
        check("PIPENET SDF 동시 생성", sdf.get("ok"),
              f"{sdf.get('bytes', 0):,}B · 노즐 {sdf.get('nozzles')}"
              f" · SLF {sdf.get('slf')}")

        r = c.get(f"/api/module-f/download?sid={sid2}&what=kfp")
        ok_dl = r.status_code == 200 and len(r.data) > 1000
        check("내려받기 .kfp", ok_dl, f"HTTP {r.status_code} · {len(r.data):,}B")
        if ok_dl:
            try:
                kfp = json.loads(r.data.decode("utf-8"))
                check("KFP 파싱", "pipe_data" in kfp and "nodes_meta_runtime" in kfp,
                      f"키 {len(kfp)}개")
            except Exception as exc:  # noqa: BLE001
                check("KFP 파싱", False, str(exc))
        r = c.get(f"/api/module-f/download?sid={sid2}&what=sdf")
        check("내려받기 .sdf", r.status_code == 200 and r.data[:5] == b"<?xml",
              f"HTTP {r.status_code} · {len(r.data):,}B")
        r = c.get(f"/api/module-f/download?sid={sid2}&what=set")
        import io as _io
        import zipfile as _zip
        if check("내려받기 한 벌(zip)", r.status_code == 200,
                 f"HTTP {r.status_code} · {len(r.data):,}B"):
            names = _zip.ZipFile(_io.BytesIO(r.data)).namelist()
            exts = sorted({os.path.splitext(n)[1] for n in names})
            check("한 벌 구성", {".kfp", ".sdf", ".slf"} <= set(exts), str(exts))

        print("\n[4-A] 최불리 30 범위로 변환")
        r = c.post("/api/module-f/convert/run",
                   json={"sid": sid2, "dto": {}, "remote_only": True})
        if check("범위 제한 변환 수락", r.get_json().get("ok")):
            if wait(c, sid2, "Remote 30 변환")["state"] == "done":
                rr = c.get(f"/api/module-f/convert/result?sid={sid2}").get_json()["result"]
                if check("범위 제한 변환 성공", rr.get("ok"),
                         json.dumps(rr.get("blockers"), ensure_ascii=False)[:200]):
                    s2 = rr["summary"]
                    check("헤드 30개로 좁혀짐", s2["heads"] == 30 and s2["remote_only"],
                          f"헤드 {s2['heads']} · 노드 {s2['nodes']} · 배관 {s2['pipes']}")
                    check("전량 대비 작아짐", s2["pipes"] < s["pipes"],
                          f"배관 {s2['pipes']} < 전량 {s['pipes']}")

        print("\n[5] Qt 미사용 확인")
        qt = [m for m in sys.modules if m.startswith("PySide6")]
        check("PySide6 미로드", not qt, str(qt))


if __name__ == "__main__":
    main()
    print("\n" + "=" * 60)
    if FAILS:
        print(f"실패 {len(FAILS)}건")
        for f in FAILS:
            print("  -", f)
        raise SystemExit(1)
    print("모듈 F 전 경로 통과")
