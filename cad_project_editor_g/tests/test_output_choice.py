# -*- coding: utf-8 -*-
"""[G17] 산출물 선택 — `.kfp` 저장을 강제로 통과시키지 않는다.

세 조합(kfp만 / sdf만 / 둘다)에서 **불리는 함수가 정확히 그것뿐인지** 본다.
「둘 다」의 동작이 이번 변경 전과 완전히 같아야 회귀가 없는 것이다(§G17).

    QT_QPA_PLATFORM=offscreen python tests/test_output_choice.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


class _Dlg:
    """변환 창 자리 — 흐름이 보는 것은 `result` 뿐이다."""

    def __init__(self, outputs):
        self.result = {"ok": True, "reason": "converted", "path": None,
                       "kfp": {"pipe_data": {}}, "outputs": outputs}


class _OldDlg:
    """`outputs` 를 안 싣는 옛 호출부 — 종전대로 둘 다 해야 한다."""

    def __init__(self):
        self.result = {"ok": True, "reason": "converted", "path": None,
                       "kfp": {"pipe_data": {}}}


def branch_calls():
    print("\n[G17] 세 조합에서 불리는 것이 그것뿐인가")
    from ui.controllers.cad_import_flow import CadImportFlow

    flow = CadImportFlow.__new__(CadImportFlow)     # __init__ 없이 분기만 본다
    called: list[str] = []
    flow._save_converted_kfp = lambda *a, **k: called.append("kfp")
    flow._open_design_input = lambda *a, **k: called.append("sdf")

    def run(dlg):
        called.clear()
        flow._after_convert(None, None, dlg)
        return list(called)

    check(".kfp 만 고르면 저장만 한다",
          run(_Dlg({"kfp": True, "sdf": False})) == ["kfp"],
          str(run(_Dlg({"kfp": True, "sdf": False}))))
    check(".sdf 만 고르면 수리계산 입력 창만 뜬다",
          run(_Dlg({"kfp": False, "sdf": True})) == ["sdf"],
          str(run(_Dlg({"kfp": False, "sdf": True}))))
    # ★순서까지 본다 — .kfp 를 저장한 «뒤» 창이 떠야 서로 영향을 안 준다(§G7).
    check("둘 다면 종전 순서(kfp → sdf) 그대로",
          run(_Dlg({"kfp": True, "sdf": True})) == ["kfp", "sdf"],
          str(run(_Dlg({"kfp": True, "sdf": True}))))
    check("outputs 가 없는 옛 호출부는 종전대로 둘 다",
          run(_OldDlg()) == ["kfp", "sdf"], str(run(_OldDlg())))
    check("둘 다 해제면 아무것도 안 부른다(창이 앞에서 막는다)",
          run(_Dlg({"kfp": False, "sdf": False})) == [],
          str(run(_Dlg({"kfp": False, "sdf": False}))))
    return True


def dialog_guard():
    print("\n[G17] 창이 «하나도 안 고름» 을 막는가")
    from PySide6.QtWidgets import QApplication
    import ui.dialogs.dialog_kfp_convert as M

    _app = QApplication.instance() or QApplication([])
    seen: list[str] = []

    class _MB:
        @staticmethod
        def warning(*a, **k):
            seen.append(str(a[2]) if len(a) > 2 else "")
            return None

        @staticmethod
        def information(*a, **k):
            return None

    dlg = M.KfpConvertDialog(payload=None)
    check("산출물 칸이 둘 다 기본 켜짐",
          dlg.chk_out_kfp.isChecked() and dlg.chk_out_sdf.isChecked(),
          f"kfp={dlg.chk_out_kfp.isChecked()} sdf={dlg.chk_out_sdf.isChecked()}")

    dlg.chk_out_kfp.setChecked(False)
    dlg.chk_out_sdf.setChecked(False)
    old_mb, M.QMessageBox = M.QMessageBox, _MB
    try:
        dlg._on_convert()
    finally:
        M.QMessageBox = old_mb
    check("둘 다 해제면 조용히 끝내지 않고 막는다",
          bool(seen) and "하나도" in seen[0], (seen[0][:44] if seen else "무반응"))
    check("막혔으면 결과가 «변환됨» 이 아니다",
          not (dlg.result or {}).get("ok"), str((dlg.result or {}).get("reason")))

    # 고른 값은 이 실행 동안 기억된다 — 매번 다시 고르지 않게.
    dlg.chk_out_kfp.setChecked(True)
    dlg.chk_out_sdf.setChecked(False)
    M.OUTPUT_CHOICE.update(dlg._outputs())
    dlg2 = M.KfpConvertDialog(payload=None)
    check("다음에 열면 같은 값으로 뜬다",
          dlg2.chk_out_kfp.isChecked() and not dlg2.chk_out_sdf.isChecked(),
          f"kfp={dlg2.chk_out_kfp.isChecked()} sdf={dlg2.chk_out_sdf.isChecked()}")
    M.OUTPUT_CHOICE.update({"kfp": True, "sdf": True})
    return True


def main() -> int:
    branch_calls()
    dialog_guard()
    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("G17 수용 기준 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
