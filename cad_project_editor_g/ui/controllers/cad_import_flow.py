"""CAD 임포트 창 흐름. 본체 파일 컨트롤러에서 찍기→손질→변환만 추출."""
from __future__ import annotations

import json
import os

from PySide6.QtCore import QThread, Qt
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QMessageBox, QProgressDialog, QWidget,
)

from services.cad_import.dto import default_out_path
from services.i18n_service import _t


class _CadEditBuildThread(QThread):
    """손질 기준망 계산만 UI 스레드 밖에서 실행한다."""

    def __init__(self, job, parent=None):
        super().__init__(parent)
        self._job = job
        self.result_value = None
        self.error = None

    def run(self):
        try:
            self.result_value = self._job()
        except BaseException as exc:
            self.error = exc


def _begin_wait_cursor() -> None:
    app = QApplication.instance()
    if app is None:
        return
    QApplication.setOverrideCursor(Qt.WaitCursor)
    app.processEvents()


def _restore_wait_cursor() -> None:
    if QApplication.instance() is None:
        return
    while QApplication.overrideCursor() is not None:
        QApplication.restoreOverrideCursor()


def _bind_cad_advance(dlg, proceed):
    """다음 임포트 창이 떠 있는 동안 현재 창을 닫지 않는다."""
    advanced = False

    def advance():
        nonlocal advanced
        advanced = True
        if proceed():
            dlg.accept()

    dlg.advance = advance
    return lambda: advanced


def _hide_cad_import_dialogs() -> None:
    from ui.dialogs.dialog_cad_edit import CadEditDialog
    from ui.dialogs.dialog_cad_pick import CadPickDialog
    from ui.dialogs.dialog_kfp_convert import KfpConvertDialog
    kinds = (CadPickDialog, CadEditDialog, KfpConvertDialog)
    for w in QApplication.topLevelWidgets():
        if isinstance(w, kinds):
            w.hide()


class CadImportFlow:
    """파일 고르기 → 찍기 → 손질 → 변환 → .kfp 저장. 본체 프로젝트는 열지 않는다."""

    def __init__(self, mw: QWidget):
        self.mw = mw

    def on_import_cad_dxf(self, path: str | None = None):
        mw = self.mw
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                mw, _t("캐드(DXF) 불러오기"), "",
                "DXF (*.dxf);;All Files (*)",
            )
        if not path:
            return
        pick_session = None
        out_dir = None
        while True:
            from ui.dialogs.dialog_cad_pick import CadPickDialog
            _begin_wait_cursor()
            try:
                dlg = CadPickDialog(
                    parent=mw, dxf_path=path, out_dir=out_dir,
                    session=pick_session)
            except ValueError as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("캐드(DXF) 불러오기"), _t(str(e)))
                return
            except Exception as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("캐드(DXF) 불러오기"), str(e))
                return
            _restore_wait_cursor()
            advanced = _bind_cad_advance(
                dlg,
                lambda: self._open_cad_edit(
                    mw, dlg.session.key, dlg.out_dir))
            if dlg.exec() != QDialog.Accepted:
                return
            if advanced():
                return
            pick_session = dlg.session
            out_dir = dlg.out_dir
            if self._open_cad_edit(mw, pick_session.key, out_dir):
                return

    def _build_cad_edit_session(self, mw, job):
        if QApplication.instance() is None or not isinstance(mw, QWidget):
            return job()

        _restore_wait_cursor()
        progress = QProgressDialog(
            _t("배관망을 구성하는 중입니다…"), "", 0, 0, mw)
        progress.setWindowTitle(_t("cad.edit.title"))
        progress.setCancelButton(None)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setWindowFlag(Qt.WindowCloseButtonHint, False)

        thread = _CadEditBuildThread(job, progress)
        thread.finished.connect(progress.accept)
        thread.start()
        progress.exec()
        thread.wait()
        if thread.error is not None:
            raise thread.error
        return thread.result_value

    def _open_cad_edit(self, mw, key, out_dir, session=None):
        from services.cad_import.edit import EditSession
        from ui.dialogs.dialog_cad_edit import CadEditDialog
        if session is None:
            _begin_wait_cursor()
            try:
                session = self._build_cad_edit_session(
                    mw,
                    lambda: EditSession.open(
                        key, out_dir=out_dir, load_saved=False,
                        use_cache=False),
                )
            except ValueError as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("cad.edit.title"), _t(str(e)))
                return True
            except SystemExit as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("cad.edit.title"), str(e))
                return True
            except Exception as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("cad.edit.title"), str(e))
                return True
        while True:
            if QApplication.instance() is None or QApplication.overrideCursor() is None:
                _begin_wait_cursor()
            try:
                dlg = CadEditDialog(parent=mw, key=key, out_dir=out_dir,
                                    session=session)
            except ValueError as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("cad.edit.title"), _t(str(e)))
                return True
            except Exception as e:
                _restore_wait_cursor()
                QMessageBox.warning(mw, _t("cad.edit.title"), str(e))
                return True
            _restore_wait_cursor()
            advanced = _bind_cad_advance(
                dlg, lambda: self._open_cad_convert(mw, dlg.session))
            result = dlg.exec()
            if result == QDialog.Accepted:
                if advanced():
                    return True
                session = dlg.session
                if self._open_cad_convert(mw, session):
                    return True
                continue
            if getattr(dlg, "went_back", False):
                return False
            return True

    def _open_cad_convert(self, mw, session):
        from ui.dialogs.dialog_kfp_convert import KfpConvertDialog
        _begin_wait_cursor()
        try:
            kwargs = {"parent": mw, "payload": session.convert_payload()}
            board = getattr(session, "board", None)
            if board is not None:
                kwargs["multi_heads"] = board.multi_arm_heads
            dlg = KfpConvertDialog(**kwargs)
        except ValueError as e:
            _restore_wait_cursor()
            QMessageBox.warning(mw, _t("cad.convert.title"), _t(str(e)))
            return True
        except Exception as e:
            _restore_wait_cursor()
            QMessageBox.warning(mw, _t("cad.convert.title"), str(e))
            return True
        _restore_wait_cursor()
        result = dlg.exec()
        if result == QDialog.Accepted:
            self._after_convert(mw, session, dlg)
            return True
        if result == QDialog.Rejected:
            return False
        return True

    def _after_convert(self, mw, session, dlg):
        """[G17] 고른 산출물만 만든다.

        종전에는 `.kfp` 저장 대화상자와 완료 알림을 **무조건** 통과해야 수리계산
        입력 창이 떴다 — SDF 만 필요한 사람에게는 불필요한 문이었다.
        `outputs` 가 없으면(옛 호출부) 종전대로 둘 다 한다.

        분기를 여기 따로 둔 것은 검사가 「고른 것만 불린다」를 확인할 이음매가
        필요해서다. 창을 띄우지 않고 이 함수만 부르면 된다.
        """
        want = (getattr(dlg, "result", None) or {}).get("outputs")             or {"kfp": True, "sdf": True}
        if want.get("kfp"):
            self._save_converted_kfp(mw, session, dlg.result)
        if want.get("sdf"):
            # [G7] `.kfp` 를 저장한 «뒤» 네 번째 창을 연다. 순서가 중요하다 —
            # 수리계산 입력 창이 떠 있어도 `.kfp` 저장은 이미 끝나 있어야
            # 서로 영향을 주지 않는다(§G7 수용 기준). 이 순서는 그대로 둔다.
            self._open_design_input(mw, session, dlg)

    def _open_design_input(self, mw, session, convert_dlg=None):
        """[G7] 수리계산 입력(SDF) 창. 실패해도 앞 단계를 무르지 않는다.

        `.kfp` 는 솔버가 전체망에서 설계구역을 스스로 고르고, `.sdf` 는 G 가
        앵커 방식으로 미리 고른다 — 둘이 다를 수 있고 그것은 버그가 아니다(§T2).
        """
        try:
            from ui.dialogs.dialog_design_input import DesignInputDialog
        except Exception as exc:      # noqa: BLE001
            print(f"[G7] 수리계산 입력 창을 열지 못했습니다: {exc}")
            return
        payload = getattr(convert_dlg, "payload", None)
        if payload is None:
            try:
                payload = session.convert_payload()
            except Exception as exc:  # noqa: BLE001
                QMessageBox.warning(mw, "수리계산 입력", str(exc))
                return
        sel = None
        res = getattr(convert_dlg, "result", None) if convert_dlg else None
        if isinstance(res, dict):
            sel = res.get("selected_source")
        try:
            dlg = DesignInputDialog(mw, session=session, payload=payload,
                                    selected_source=sel)
        except Exception as exc:      # noqa: BLE001
            QMessageBox.warning(mw, "수리계산 입력", str(exc))
            return
        # 직전 K·선정 결과는 창 객체가 들고 있다 — 세션에 얹어 다시 열 때 잇는다.
        prev = getattr(session, "_design_dialog_state", None)
        if isinstance(prev, dict) and prev.get("k"):
            dlg.spin_k.setValue(int(prev["k"]))
        dlg.exec()
        try:
            session._design_dialog_state = {"k": int(dlg.spin_k.value())}
        except Exception:             # noqa: BLE001
            pass

    def _save_converted_kfp(self, mw, session, result):
        _hide_cad_import_dialogs()
        kfp = (result or {}).get("kfp")
        path = (result or {}).get("path")
        if not kfp and not path:
            return
        if not path:
            suggested = default_out_path(session.key)
            path, _ = QFileDialog.getSaveFileName(
                mw, _t("cad.convert.save_title"), suggested,
                "KFP (*.kfp);;All Files (*)",
            )
            if not path:
                return
            try:
                os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(kfp, f, ensure_ascii=False, indent=2)
            except Exception as e:
                QMessageBox.warning(mw, _t("cad.convert.title"), str(e))
                return
        QMessageBox.information(
            mw, _t("cad.convert.title"),
            _t("cad.convert.saved", path=path))
