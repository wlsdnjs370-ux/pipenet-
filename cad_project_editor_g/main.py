"""CAD 임포트 독립 실행기. DXF → 찍기 → 손질 → 변환 → .kfp. 수리계산은 없다."""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from PySide6.QtWidgets import (
    QApplication, QLabel, QMainWindow, QPushButton, QVBoxLayout, QWidget,
)

from services.i18n_service import I18nService, _t
from ui.controllers.cad_import_flow import CadImportFlow


def _ensure_head_library(mw: QMainWindow) -> None:
    """변환 창의 K 목록이 parent.editor.get_head_items 를 쓰므로 라이브러리만 채운다."""
    if getattr(mw, "editor", None) is not None:
        return
    from editor_core import PipeEditor
    from services.pipenet_import import _ensure_default_libraries
    editor = PipeEditor()
    _ensure_default_libraries(editor)
    mw.editor = editor


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.editor = None
        self.setWindowTitle("CAD 임포트")
        self.resize(520, 280)
        self.flow = CadImportFlow(self)

        hint = QLabel(
            "DXF를 고른 뒤 배관·헤드를 찍고, 손질한 다음 .kfp로 저장합니다.\n"
            "수리계산은 이 프로그램에 없습니다."
        )
        hint.setWordWrap(True)
        btn = QPushButton(_t("캐드(DXF) 불러오기"))
        btn.clicked.connect(self._on_open_dxf)

        root = QWidget()
        lay = QVBoxLayout(root)
        lay.addWidget(hint)
        lay.addWidget(btn)
        lay.addStretch(1)
        self.setCentralWidget(root)

        open_act = self.menuBar().addMenu("파일").addAction(_t("캐드(DXF) 불러오기"))
        open_act.triggered.connect(self._on_open_dxf)

    def _on_open_dxf(self):
        _ensure_head_library(self)
        self.flow.on_import_cad_dxf()


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv if argv is None else argv)
    I18nService.instance().load("ko")
    app = QApplication.instance() or QApplication(argv)
    win = MainWindow()
    win.show()
    if "--self-check" in argv:
        print("OK", win.windowTitle(), flush=True)
        return 0
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
