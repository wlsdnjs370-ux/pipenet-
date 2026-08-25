# -*- coding: utf-8 -*-
"""[G7] 수리계산 입력 창 — 실제 Qt 위젯으로 띄워 확인한다.

창은 헤드리스 검사로 못 잡는 것이 있다(위젯이 실제로 붙었나, 막힘 사유가
화면에 돌아오나, 제외 헤드가 눈에 보이나). offscreen 플랫폼으로 띄워
사람 화면에는 아무것도 남기지 않는다.

    python tests/test_design_dialog.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)
# 사람 화면에 창을 띄우지 않는다 — 검사가 데스크톱을 어지럽히면 안 된다.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

KEY = "B1F 현장조사 소화설비 평면도"
OUT_DIR = _ROOT / "tests" / "_out"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def main() -> int:
    from PySide6.QtWidgets import QApplication
    from services.cad_import.edit.session import EditSession
    from ui.dialogs.dialog_design_input import DesignInputDialog

    app = QApplication.instance() or QApplication([])
    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None

    print("\n[G7] 수리계산 입력 창")
    dlg = DesignInputDialog(None, session=es, payload=payload,
                            selected_source=sel, k=30)
    # ★창을 띄워야 isVisible() 이 뜻을 갖는다. 부모 창이 숨어 있으면 자식은
    #   show() 해도 isVisible()==False 라, 안 띄우고 재면 «안 보인다» 로 잘못
    #   읽힌다(실제로 그렇게 헛돌았다). offscreen 이라 사람 화면엔 안 뜬다.
    dlg.show()
    app.processEvents()
    check("창 생성", dlg is not None, dlg.windowTitle())
    check("기준개수 K 기본 30", dlg.spin_k.value() == 30, str(dlg.spin_k.value()))
    check("저장 단추는 계산 전 잠김", not dlg.btn_save.isEnabled())

    # ── 막힘 경로 — 급수원이 없으면 사유를 화면에 돌려준다(조용히 실패 금지)
    keep = list(es.board.sources)
    es.board.sources = []
    dlg._on_run()
    app.processEvents()
    blocked_msg = dlg.lbl_warn.text()
    check("급수원 없으면 막고 사유를 보여준다",
          dlg.lbl_warn.isVisible() and "급수" in blocked_msg,
          blocked_msg[:46])
    check("막혔을 때 저장이 열리지 않는다", not dlg.btn_save.isEnabled())
    es.board.sources = keep

    # ── 정상 경로
    dlg._on_run()
    app.processEvents()
    check("계산 성공 후 저장 열림", dlg.btn_save.isEnabled(),
          dlg.lbl_warn.text()[:60])
    summary = dlg.lbl_sum.text()
    for want in ("설계면적", "앵커", "corridor 총연장", "관경 근거", "부속"):
        if not check(f"요약에 «{want}»", want in summary, ""):
            break
    print("      " + summary.replace("\n", " · ")[:150])

    # ★제외 헤드가 반드시 눈에 보여야 한다(BLOCKED B4)
    warn = dlg.lbl_warn.text()
    got = dlg.result.get("worst") or {}
    excluded = int(got.get("excluded_heads") or 0)
    check("제외 헤드 수를 화면에 보여준다",
          (not excluded) or (dlg.lbl_warn.isVisible() and "제외" in warn),
          f"제외 {excluded:,}개 · 경고 «{warn[:52]}»")
    if excluded:
        check("많이 빠지면 배관을 이으라고 알린다", "이어" in warn,
              warn[-46:])

    # ── 창을 닫았다 다시 열어도 K 가 유지되는가(§G7 수용 기준)
    dlg.spin_k.setValue(20)
    state = {"k": int(dlg.spin_k.value())}
    dlg2 = DesignInputDialog(None, session=es, payload=payload,
                             selected_source=sel, k=state["k"])
    check("다시 열어도 직전 K 유지", dlg2.spin_k.value() == 20,
          str(dlg2.spin_k.value()))

    from services.cad_import.design.emit import emit_design_sdf

    # ── [G16] 미리보기 — 저장하기 전에 형태와 표 값을 창 안에서 본다
    import time
    import xml.etree.ElementTree as ET

    dlg.chk_iso.setChecked(True)
    dlg._redraw()
    pts = dict(dlg._preview_points)
    check("미리보기가 좌표를 그렸다", len(pts) > 0, f"노드 {len(pts)}개")
    items = len(dlg.view_iso.scene().items())
    check("도면에 그린 것이 있다", items > 0, f"{items}개")

    # ★핵심 — 화면에 그린 좌표와 파일에 들어간 좌표가 같아야 미리보기다.
    prev = emit_design_sdf(dlg.result["tables"], OUT_DIR / "g16_preview.sdf",
                           project_title="G16 미리보기 대조", **dlg._view_opts())
    r = ET.parse(str(prev)).getroot()
    saved = {}
    for n in r.iter("Node"):
        q = n.find("Position")
        if q is not None:
            saved[str(n.get("label"))] = (float(q.get("x")), float(q.get("y")))
    common = set(pts) & set(saved)
    worst = max((max(abs(pts[k][0] - saved[k][0]), abs(pts[k][1] - saved[k][1]))
                 for k in common), default=0.0)
    # ★writer 는 좌표를 `.6g`(유효 6자리)로 찍는다. 그래서 절대오차 1e-6 비교는
    #   애초에 설 수 없고, 남은 차이는 전부 그 자리수 자르기다. 같은 형식으로
    #   찍어 **문자열이 일치**하는지 보면 두 수가 같은 데서 나왔음이 증명된다.
    fmt = lambda v: format(float(v), ".6g")
    mism = [k for k in common
            if (fmt(pts[k][0]), fmt(pts[k][1])) != (fmt(saved[k][0]), fmt(saved[k][1]))]
    check("미리보기 좌표 == 저장된 Position (writer 자리수로)",
          len(common) == len(saved) and not mism,
          f"노드 {len(common)}/{len(saved)} · 어긋난 것 {len(mism)}개 · "
          f"자리수 자르기로 인한 최대 차이 {worst:.3e}")

    # 표시 설정을 바꾸면 그림은 바뀌고 표 값은 안 바뀐다(표시 전용 증명).
    def _pipe_view():
        dlg.cmb_table.setCurrentIndex(1)
        dlg._on_table_switch()
        return [[dlg.tbl.item(i, j).text() if dlg.tbl.item(i, j) else ""
                 for j in range(dlg.tbl.columnCount())]
                for i in range(dlg.tbl.rowCount())]

    before_tbl = _pipe_view()
    t0 = time.time()
    dlg.spin_canvas.setValue(6000)         # 신호가 _redraw 를 부른다
    elapsed = time.time() - t0
    after = dict(dlg._preview_points)
    after_tbl = _pipe_view()
    moved = sum(1 for k in pts if k in after and after[k] != pts[k])
    check("보기 설정을 바꾸면 도면이 즉시 바뀐다", moved > 0,
          f"{moved}개 노드가 움직임")
    check("그래도 표 값은 그대로", before_tbl == after_tbl,
          "달라짐" if before_tbl != after_tbl else f"{len(after_tbl)}행 동일")
    check("다시 그리기가 0.5초 안", elapsed < 0.5, f"{elapsed*1000:.0f} ms")
    dlg.spin_canvas.setValue(3000)

    # ★「고도 펼침 배율」은 이 도면에서 그림을 못 바꾼다 — 노드 표고가 **전부 0** 이라
    #   lift 가 0 이기 때문이다(단층 평면도). 기구가 죽은 것이 아니라 펼칠 높이가
    #   없는 것이고, 그 증명은 합성망 검사(test_sdf_post g12)가 이미 하고 있다.
    elevs = {float(n.get("elevation", 0) or 0)
             for n in dlg.result["tables"].nodes}
    zs_before = dict(dlg._preview_points)
    dlg.spin_zscale.setValue(2.0)
    zs_moved = sum(1 for k in zs_before
                   if dlg._preview_points.get(k) != zs_before[k])
    check("표고가 없으면 고도 배율이 그림을 안 바꾼다(정직한 무변화)",
          (len(elevs) == 1 and zs_moved == 0) or zs_moved > 0,
          f"표고 종류 {sorted(elevs)} · 움직인 노드 {zs_moved}개")
    dlg.spin_zscale.setValue(1.0)

    # 표는 저장될 값 그대로 — 노드/배관/노즐/부속 넷이 다 보인다.
    kinds = [dlg.cmb_table.itemData(i) for i in range(dlg.cmb_table.count())]
    check("표 네 가지를 다 보여 준다",
          kinds == ["nodes", "pipes", "nozzles", "fittings"], str(kinds))
    dlg.cmb_table.setCurrentIndex(1)
    dlg._on_table_switch()
    hdr = [dlg.tbl.horizontalHeaderItem(j).text()
           for j in range(dlg.tbl.columnCount())]
    check("배관표에 길이·관경이 있다",
          "길이(m)" in hdr and "호칭경(mm)" in hdr, str(hdr))
    # ★관경 근거는 요약의 집계로만 있었다 — «이 배관» 이 도면 텍스트에서 온 것인지
    #   별표1 폴백인지 행에서 확인할 수 있어야 한다(§G16).
    check("배관표에 관경 근거 열이 있다", "관경 근거" in hdr, str(hdr))
    j = hdr.index("관경 근거") if "관경 근거" in hdr else -1
    vals = ({dlg.tbl.item(i, j).text() for i in range(dlg.tbl.rowCount())
             if dlg.tbl.item(i, j)} if j >= 0 else set())
    check("근거가 사람 말로 채워진다",
          bool(vals) and all(v and "nfpc" not in v for v in vals), str(sorted(vals)))
    check("배관표 행 수가 표와 같다",
          dlg.tbl.rowCount() == len(dlg.result["tables"].pipes),
          f"{dlg.tbl.rowCount()} / {len(dlg.result['tables'].pipes)}")

    # 행을 고르면 도면에서 강조된다.
    if dlg.tbl.rowCount():
        dlg.tbl.selectRow(0)
        lab = dlg.tbl.item(0, 0).text()
        it, pen = dlg._pipe_items.get(lab, (None, None))
        check("행을 고르면 그 배관이 굵어진다",
              it is not None and it.pen().widthF() > pen.widthF(),
              f"{lab} · {pen.widthF():.1f} → {it.pen().widthF():.1f}"
              if it is not None else "그 배관을 못 찾음")

    # 평면 보기에서는 헤드 스텁을 적용하지 않는다(§G15 수용 기준).
    dlg.chk_iso.setChecked(False)
    flat = dict(dlg._preview_points)
    dlg.chk_iso.setChecked(True)
    check("아이소를 끄면 좌표가 평면으로 돌아온다",
          flat != dict(dlg._preview_points), "같으면 표시가 안 바뀐 것")

    # 창을 키워도 레이아웃이 버틴다.
    dlg.resize(1920, 1080)
    QApplication.processEvents()
    check("최대화해도 미리보기가 살아 있다",
          dlg.view_iso.width() > 0 and dlg.view_iso.height() > 0,
          f"{dlg.view_iso.width()}×{dlg.view_iso.height()}")


    # ── 저장 (파일 대화상자 없이 직접) — SDF+SLF 한 쌍
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = emit_design_sdf(dlg.result["tables"], OUT_DIR / "g7_dialog.sdf",
                          project_title="G7 창 검증")
    check("창 결과로 SDF+SLF 저장", out.is_file()
          and out.with_suffix(".slf").is_file(),
          f"{out.stat().st_size:,} B + {out.with_suffix('.slf').stat().st_size:,} B")

    dlg.close()
    dlg2.close()
    app.processEvents()

    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("G7 수용 기준 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
