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
            self._save_converted_kfp(mw, session, dlg.result)
            return True
        if result == QDialog.Rejected:
            return False
        return True

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
