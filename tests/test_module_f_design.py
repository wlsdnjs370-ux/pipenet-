# -*- coding: utf-8 -*-
"""[F-2] design/ HTTP 노출 검증.

수용 기준 세 가지:
  ① 같은 손질 저장본·같은 설정으로 **G 데스크톱(4번째 창)** 과 F 웹이 만든
     `.sdf` 가 완전히 동일하다(diff 0) — 같은 엔진이라는 증명.
  ② K 를 바꾸면 build 만 다시 돌고, iso 설정만 바꾸면 최불리 재계산 없이
     preview 만 갱신된다(캐시).
  ③ preview 좌표 == 저장된 SDF 의 Position (writer 자리수 `.6g` 로 비교).

①은 엔진 함수를 직접 부르는 반쪽 증명이 아니라 **Qt 대화상자를 offscreen 으로
실제로 띄워** 그 저장 경로와 견준다 — 대화상자가 엔진과 같다는 것은
`cad_project_editor_g/tests/test_design_dialog.py` 가, 웹이 대화상자와 같다는
것은 여기가 못박는다.

    QT_QPA_PLATFORM=offscreen python tests/test_module_f_design.py
"""
from __future__ import annotations

import importlib.util
import os
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEY = "B1F 현장조사 소화설비 평면도"
FAILS: list[str] = []
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


# ★대기 상한 — B1F(간선 23,273)의 «표 확정» 은 실측 200초대다. 종전 600회
#   ×0.3s = 180초로는 머신이 조금만 바빠도 넘어가, 결함이 아닌데 붉게 뜬다
#   (실측: 200.3s 에서 preview 가 「아직 계산 중」 → KeyError 'view').
#   재는 것은 «라우트가 다 있는가» 이지 속도가 아니므로 넉넉히 둔다.
def _wait(c, sid, limit=1400):
    for _ in range(limit):
        jb = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if jb.get("state") in ("done", "error", "idle"):
            return jb
        time.sleep(0.3)
    return jb


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
        jb = _wait(c, sid)
        if not check("B1F reopen", jb.get("state") == "done", str(jb)[:70]):
            return 1

        print("\n[표 확정 — design/build]")
        r = c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        check("build 수락", r.get_json().get("ok"), str(r.get_json())[:70])
        jb = _wait(c, sid)
        check("build 완료(잡 규약)", jb.get("state") == "done"
              and jb.get("phase") == "수리계산 입력",
              f"{jb.get('phase')} · {jb.get('elapsed')}s")
        r = c.get(f"/api/module-f/design/preview?sid={sid}&iso=1")
        pv = r.get_json()
        check("preview 가 온다", pv.get("ok"),
              str(pv)[:70] if not pv.get("ok") else
              f"노드 {len(pv['view']['nodes'])} · 배관 {len(pv['view']['pipes'])}")
        check("표 4종이 저장될 값 그대로", all(
            k in (pv.get("tables") or {}) for k in
            ("nodes", "pipes", "nozzles", "fittings")),
            str(list((pv.get("tables") or {}).keys())))
        check("배관 표에 관경 근거가 있다",
              all("src" in p for p in pv["view"]["pipes"]),
              f"{pv['view']['pipes'][0] if pv['view']['pipes'] else '없음'}"[:60])

        print("\n[캐시 — 보기 설정은 build 를 다시 돌리지 않는다]")
        t0 = time.time()
        r = c.get(f"/api/module-f/design/preview?sid={sid}&iso=1"
                  f"&iso_z_scale=2.0")
        pv2 = r.get_json()
        dt = time.time() - t0
        jb2 = c.get(f"/api/module-f/job?sid={sid}").get_json()
        check("preview 만 갱신 — 잡이 새로 돌지 않음",
              pv2.get("ok") and jb2.get("phase") == "수리계산 입력"
              and jb2.get("state") == "done", f"{jb2.get('phase')}")
        check("보기 변경이 0.5초 안", dt < 0.5, f"{dt*1000:.0f} ms")
        moved = sum(1 for a, b in zip(pv["view"]["nodes"], pv2["view"]["nodes"])
                    if (a["x"], a["y"]) != (b["x"], b["y"]))
        same_tbl = pv["tables"]["pipes"] == pv2["tables"]["pipes"]
        # 이 도면은 표고가 전부 0 이라 zscale 은 그림을 못 바꾼다 — 캔버스로 본다.
        r = c.get(f"/api/module-f/design/preview?sid={sid}&canvas_units=6000")
        pv3 = r.get_json()
        moved3 = sum(1 for a, b in zip(pv["view"]["nodes"], pv3["view"]["nodes"])
                     if (a["x"], a["y"]) != (b["x"], b["y"]))
        check("보기 설정이 좌표를 실제로 바꾼다(캔버스)", moved3 > 0,
              f"{moved3}개 노드")
        check("그래도 표 값은 그대로", same_tbl
              and pv["tables"]["pipes"] == pv3["tables"]["pipes"], "표 불변")
        c.get(f"/api/module-f/design/preview?sid={sid}&canvas_units=3000")

        print("\n[emit — .sdf + .slf 한 쌍]")
        r = c.post("/api/module-f/design/emit",
                   json={"sid": sid, "iso": True, "iso_z_scale": 1.0,
                         "canvas_units": 3000, "head_stub_pct": 2.5,
                         "lift_ref": "valve"})
        em = r.get_json()
        if not check("emit 성공", em.get("ok"), str(em)[:90]):
            return 1
        # ★**이번 세션** 것을 집는다. 종전에는 `glob("*_design")` 을 돌며 마지막
        #   것을 남겼는데, 그 폴더에는 지난 실행의 산출물이 쌓여 있어 glob 순서에
        #   따라 **옛 파일**을 집었다. 실측으로 그렇게 배관 6개짜리 옛 SDF 를
        #   집어 「웹 6개 vs 데스크톱 110개」로 실패했다 — 제품이 아니라 이 시험이
        #   틀린 것이었다(같은 시각 실제 산출물은 110개였다).
        web_sdf = None
        cand = (ROOT / "data" / "uploads" / "module_f" / f"{sid}_design"
                / f"{KEY}_수리계산입력.sdf")
        if cand.is_file():
            web_sdf = cand
        else:
            # 세션 폴더 이름 규약이 바뀌었으면 «가장 최근» 으로 물러서되,
            # 조용히 아무거나 집지는 않는다.
            pool = sorted((ROOT / "data" / "uploads" / "module_f").glob(
                f"*_design/{KEY}_수리계산입력.sdf"),
                key=lambda p: p.stat().st_mtime, reverse=True)
            web_sdf = pool[0] if pool else None
            if web_sdf is not None:
                print(f"      (세션 폴더를 못 찾아 최신 산출물로 대체: "
                      f"{web_sdf.parent.name})")
        check("SDF+SLF 가 실제로 있다", web_sdf is not None
              and web_sdf.with_suffix(".slf").is_file(),
              str(em.get("sdf")))
        r = c.get(f"/api/module-f/download?sid={sid}&what=design")
        check("내려받기 — zip 한 벌", r.status_code == 200
              and r.data[:2] == b"PK", f"HTTP {r.status_code} · {len(r.data):,}B")

        print("\n[preview 좌표 == 저장된 Position]")
        rr = ET.parse(str(web_sdf)).getroot()
        saved = {}
        for n in rr.iter("Node"):
            q = n.find("Position")
            if q is not None:
                saved[str(n.get("label"))] = (float(q.get("x")),
                                              float(q.get("y")))
        r = c.get(f"/api/module-f/design/preview?sid={sid}&iso=1")
        pts = {n["label"]: (n["x"], n["y"])
               for n in r.get_json()["view"]["nodes"]}
        fmt = lambda v: format(float(v), ".6g")
        mism = [k for k in saved
                if k not in pts
                or (fmt(pts[k][0]), fmt(pts[k][1]))
                != (fmt(saved[k][0]), fmt(saved[k][1]))]
        check("좌표 일치(writer 자리수)", not mism and len(saved) > 0,
              f"노드 {len(saved)} · 어긋남 {len(mism)}")

    print("\n[G 데스크톱 4번째 창과 diff 0]")
    g_root = ROOT / "cad_project_editor_g"
    for p in (str(g_root),):
        if p in sys.path:
            sys.path.remove(p)
    import subprocess
    code = r'''
import os, sys
sys.stdout.reconfigure(errors="replace")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from pathlib import Path
sys.path.insert(0, str(Path.cwd())); sys.path.insert(0, str(Path.cwd().parent))
from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication([])
from services.cad_import.edit.session import EditSession
from ui.dialogs.dialog_design_input import DesignInputDialog
KEY = "B1F 현장조사 소화설비 평면도"
es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
payload = es.convert_payload()
srcs = payload.get("sources") or ()
sel = srcs[0].get("tag") if len(srcs) > 1 else None
dlg = DesignInputDialog(None, session=es, payload=payload, selected_source=sel, k=30)
dlg.show(); dlg._on_run()
from services.cad_import.design.emit import emit_design_sdf
out_dir = Path("tests/_out/f2_desktop"); out_dir.mkdir(parents=True, exist_ok=True)
dlg.chk_iso.setChecked(True)
dlg.spin_zscale.setValue(1.0); dlg.spin_canvas.setValue(3000)
dlg.spin_stub.setValue(2.5)
out = emit_design_sdf(dlg.result["tables"], out_dir / (KEY + "_수리계산입력.sdf"),
                      project_title=f"{KEY} 수리계산 입력", **dlg._view_opts())
print("DESKTOP_SDF=", out)
'''
    # ★자식의 출력 인코딩을 **못박는다.** 부모는 UTF-8 로 읽는데 자식은
    #   Windows 기본(cp949)으로 쓰므로, 그냥 두면 경로의 한글이 깨져
    #   `DESKTOP_SDF=` 를 읽고도 «그런 파일 없음» 이 된다.
    #   바깥 셸에 `PYTHONIOENCODING` 이 있으면 통과하고 없으면 실패했다 —
    #   즉 «직접 돌리면 되는데 pytest 로는 안 되는» 시험이었다.
    _env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, "-c", code], cwd=str(g_root),
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=1200, env=_env)
    desk_sdf = None
    for ln in (r.stdout or "").splitlines():
        if ln.startswith("DESKTOP_SDF="):
            desk_sdf = Path(ln.split("=", 1)[1].strip())
            if not desk_sdf.is_absolute():
                desk_sdf = g_root / desk_sdf     # 자식의 cwd 는 G 트리다
    if not check("G 데스크톱 창 저장 성공", desk_sdf is not None
                 and desk_sdf.is_file(),
                 f"rc={r.returncode} · DESKTOP_SDF={desk_sdf}"):
        # ★사유를 120자로 자르면 원인을 가린다 — 자식이 왜 죽었는지가 전부다.
        print("  ── 자식 stderr ──")
        print((r.stderr or "").strip()[-1500:] or "(비어 있음)")
        print("  ── 자식 stdout 끝 ──")
        print((r.stdout or "").strip()[-600:] or "(비어 있음)")
        return 1
    web = web_sdf.read_text(encoding="utf-8")
    desk = desk_sdf.read_text(encoding="utf-8")
    # ★배관 label 만은 실행마다 다르다 — 엔진의 id 부여가 set 순회 순서를 타는
    #   알려진 비결정성(G BLOCKED B7 — .kfp 도 같은 이유로 구조 지문 비교다).
    #   그래서 «Pipe label 정규화 후 바이트 동일» + «label 집합은 순열» 로 세운다.
    #   노드 label·좌표·bore·length·rise 는 전부 그대로 비교된다.
    import re
    web_labels = sorted(re.findall(r'<Pipe [^>]*label="(P\d+)"', web))
    desk_labels = sorted(re.findall(r'<Pipe [^>]*label="(P\d+)"', desk))
    check("웹 == 데스크톱 · Pipe label 은 서로 순열", web_labels == desk_labels,
          f"{len(web_labels)}개 vs {len(desk_labels)}개")
    norm = lambda t: re.sub(r'label="P\d+"', 'label="P#"', t)
    check("웹 == 데스크톱 · label 정규화 후 바이트 동일",
          norm(web) == norm(desk),
          f"웹 {len(web):,}B vs 데스크톱 {len(desk):,}B (B7 — label 만 비결정)")
    web_slf = web_sdf.with_suffix(".slf").read_bytes()
    desk_slf = desk_sdf.with_suffix(".slf").read_bytes()
    check("웹 == 데스크톱 · SLF 바이트 동일", web_slf == desk_slf,
          f"{len(web_slf):,}B")

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("F-2 design/ HTTP 노출 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())


# ─────────────────────────────────────────────────────────────────────────
def test_design_라우트가_열려_있다():
    """pytest 진입점 — `main()` 을 그대로 태운다.

    ★이 파일은 오랫동안 `test_*.py` 이면서 **pytest 수집 0건**이었다.
      `assert` 대신 `FAILS`+`return 1` 로 보고하는 스크립트였기 때문이다.
      「F-2 design/ HTTP 노출」를 표방하면서 자동 실행에는 없었다 — 안 도는 시험은
      통과가 아니라 없는 것이다.
    """
    assert main() == 0, "실패 상세는 위 출력 참고"
