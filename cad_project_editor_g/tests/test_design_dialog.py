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

    # ── 저장 (파일 대화상자 없이 직접) — SDF+SLF 한 쌍
    from services.cad_import.design.emit import emit_design_sdf
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
