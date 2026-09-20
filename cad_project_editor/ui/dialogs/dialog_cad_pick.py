# 파일명: ui/dialogs/dialog_cad_pick.py
"""배관/헤드 추출 — 화면만. 판정은 services.cad_import.pick."""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if __package__ in (None, "") and str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from PySide6.QtCore import QPoint, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QDialog, QFormLayout, QFrame,
    QGraphicsScene, QGraphicsView, QGroupBox, QHBoxLayout, QLabel,
    QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from services.cad_import.colors import C_DARK, rgb_dark
from services.cad_import.kinds import SLOT_COMBO, SLOT_UPDOWN
from services.cad_import.pick import PICK_PX, PickSession
from services.i18n_service import I18nService, _t

_PANEL_W = 300
_SAMPLE_DXF = _ROOT / "docs" / "import" / "DWG" / "3f sample_libredwg.dxf"
_DRAG_PX = 5.0


def _ensure_i18n() -> None:
    if _t("cad.pick.title") == "cad.pick.title":
        I18nService.instance().load("ko")


def _guide(key: str) -> QLabel:
    """카드 맨 위 사용법 — 단계별 ※ 안내보다 위에 있다는 뜻으로 파랑."""
    lbl = QLabel(_t(key))
    lbl.setWordWrap(True)
    lbl.setObjectName("CadPickGuide")
    lbl.setStyleSheet("QLabel#CadPickGuide { font-weight: bold; }")
    return lbl


def _hint(key: str) -> QLabel:
    """단계 안의 안내 — 회색 처리된 단계에서는 같이 흐려진다."""
    lbl = QLabel(_t(key))
    lbl.setWordWrap(True)
    lbl.setObjectName("CadPickHint")
    lbl.setStyleSheet("QLabel#CadPickHint { font-weight: bold; }")
    return lbl


# 윈도우 네이티브 체크는 어두운 배경에 어두운 글자라 안 보인다.
# 체크된 단추만 검정. 눌림(isDown)까지 칠하면 이미 켜진 단추와 같이 검정이 된다.
_CHECKED_BTN_QSS = (
    "QPushButton { color: #ffffff; background-color: #2b2b2b;"
    " border: 1px solid #2b2b2b; }"
)


def _bind_checked_contrast(*buttons: QPushButton) -> None:
    def _apply_all(*_args) -> None:
        for btn in buttons:
            btn.setStyleSheet(_CHECKED_BTN_QSS if btn.isChecked() else "")

    for btn in buttons:
        btn.toggled.connect(_apply_all)
    _apply_all()


def _mode_button(text: str, *, first: bool) -> QPushButton:
    btn = QPushButton(text)
    btn.setCheckable(True)
    btn.setProperty("segpos", "first" if first else "last")
    return btn


class CadPickCanvas(QGraphicsView):
    """다크 도면. 클릭=찍기 · 끌기=이동 · 휠=확대. 판정은 세션."""

    picked = Signal(float, float, float)

    def __init__(self, session: PickSession, parent=None):
        super().__init__(parent)
        self.setObjectName("CadPickCanvas")
        self.setFrameShape(QFrame.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setBackgroundBrush(QColor(0, 0, 0))
        self.setScene(QGraphicsScene(self))
        # 카메라는 도면 items 박스에 가두지 않는다. 맞춤은 itemsBoundingRect.
        self.setSceneRect(QRectF(-1e9, -1e9, 2e9, 2e9))
        self._session = session
        self._hl = []
        self._pipe_paths = {}
        self._user_view = False
        self._drag = {"on": False, "moved": False, "pos": QPoint()}
        self._draw_world()

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        if not self._user_view:
            self._fit()

    def wheelEvent(self, ev):
        if ev.angleDelta().y() == 0:
            return
        self._user_view = True
        k = 1.25 if ev.angleDelta().y() > 0 else 0.8
        # 거대 sceneRect 에서 AnchorUnderMouse 는 첫 휠에 원점(0,0)으로 붙는다.
        # 팬과 같은 mapToScene 차이로 마우스 아래 점을 유지한다.
        pos = ev.position().toPoint()
        self.setTransformationAnchor(QGraphicsView.NoAnchor)
        a = self.mapToScene(pos)
        self.scale(k, k)
        b = self.mapToScene(pos)
        self.translate(b.x() - a.x(), b.y() - a.y())

    def mousePressEvent(self, ev):
        if ev.button() != Qt.LeftButton:
            super().mousePressEvent(ev)
            return
        self._drag = {"on": True, "moved": False, "pos": ev.position().toPoint()}
        self.viewport().grabMouse()

    def mouseMoveEvent(self, ev):
        if not self._drag["on"]:
            super().mouseMoveEvent(ev)
            return
        pos = ev.position().toPoint()
        delta = pos - self._drag["pos"]
        if not self._drag["moved"]:
            if (delta.x() ** 2 + delta.y() ** 2) ** 0.5 < _DRAG_PX:
                return
            self._drag["moved"] = True
            self._user_view = True
        self._pan_view(self._drag["pos"], pos)
        self._drag["pos"] = pos

    def mouseReleaseEvent(self, ev):
        if ev.button() != Qt.LeftButton or not self._drag["on"]:
            super().mouseReleaseEvent(ev)
            return
        moved = self._drag["moved"]
        self._drag["on"] = False
        self.viewport().releaseMouse()
        if moved:
            return
        x, y = self._world_xy(ev.position().toPoint())
        self.picked.emit(x, y, self._pick_mm())

    def _pan_view(self, from_pos, to_pos) -> None:
        """화면 픽셀 끌기를 도면 이동으로 바꾼다. translate 인자는 scene mm."""
        self.setTransformationAnchor(QGraphicsView.NoAnchor)
        a = self.mapToScene(from_pos)
        b = self.mapToScene(to_pos)
        self.translate(b.x() - a.x(), b.y() - a.y())

    def _world_xy(self, pos) -> tuple[float, float]:
        p = self.mapToScene(pos)
        return float(p.x()), float(-p.y())

    def _pick_mm(self) -> float:
        p0 = self.mapToScene(QPoint(0, 0))
        p1 = self.mapToScene(QPoint(int(PICK_PX), 0))
        return abs(p1.x() - p0.x())

    def _fit(self) -> None:
        r = self.scene().itemsBoundingRect()
        if r.isValid() and r.width() > 0 and r.height() > 0:
            self.fitInView(r, Qt.KeepAspectRatio)

    def _draw_world(self) -> None:
        scene = self.scene()
        scene.clear()
        self._hl = []
        self._pipe_paths = {}
        w = self._session.world
        seg_paths = {}
        for _ly, c, a, b in w.segs:
            path = seg_paths.setdefault(c, QPainterPath())
            path.moveTo(a[0], -a[1])
            path.lineTo(b[0], -b[1])
        for c, path in seg_paths.items():
            pen = QPen(QColor(rgb_dark(c)))
            pen.setCosmetic(True)
            scene.addPath(path, pen)
        circle_paths = {}
        for _ly, c, x, y, r in w.circles:
            path = circle_paths.setdefault(c, QPainterPath())
            path.addEllipse(x - r, -y - r, 2 * r, 2 * r)
        for c, path in circle_paths.items():
            pen = QPen(QColor(rgb_dark(c)))
            pen.setCosmetic(True)
            scene.addPath(path, pen)
        pipe_pen = QPen(QColor(C_DARK["재료"]))
        pipe_pen.setCosmetic(True)
        pipe_pen.setWidth(2)
        head_pen = QPen(QColor(C_DARK["헤드원"]))
        head_pen.setCosmetic(True)
        head_pen.setWidth(2)
        mark_pen = QPen(QColor(C_DARK["안붙은끝점"]))
        mark_pen.setCosmetic(True)
        self._hl = [
            scene.addPath(QPainterPath(), pipe_pen),
            scene.addPath(QPainterPath(), head_pen),
            scene.addPath(QPainterPath(), head_pen),
            scene.addPath(QPainterPath(), mark_pen),
        ]
        for item in self._hl:
            item.setZValue(10.0)
        self.refresh_highlight()

    def _pipe_path(self, bundle) -> QPainterPath:
        """배관 묶음 경로를 최초 한 번만 만든다."""
        key = tuple(bundle)
        cached = self._pipe_paths.get(key)
        if cached is not None:
            return cached
        path = QPainterPath()
        for a, b in self._session.board.by_bundle.get(key, ()):
            path.moveTo(a[0], -a[1])
            path.lineTo(b[0], -b[1])
        self._pipe_paths[key] = path
        return path

    def refresh_highlight(self) -> None:
        geom = self._session.highlight_geom()
        pipe_path = QPainterPath()
        for bundle in geom.get("pipe_bundles") or ():
            pipe_path.addPath(self._pipe_path(bundle))
        head_path = QPainterPath()
        for x, y, r in geom["head_circles"]:
            head_path.addEllipse(x - r, -y - r, 2 * r, 2 * r)
        tri_path = QPainterPath()
        for a, b in geom["tri_segs"]:
            tri_path.moveTo(a[0], -a[1])
            tri_path.lineTo(b[0], -b[1])
        mark_path = QPainterPath()
        last = geom.get("last_click")
        if last:
            x, y = last
            s = 80.0
            mark_path.moveTo(x - s, -y - s)
            mark_path.lineTo(x + s, -y + s)
            mark_path.moveTo(x - s, -y + s)
            mark_path.lineTo(x + s, -y - s)
        for item, path in zip(
                self._hl, (pipe_path, head_path, tri_path, mark_path)):
            item.setPath(path)


class CadPickDialog(QDialog):
    def __init__(self, parent=None, dxf_path=None, out_dir=None, session=None):
        super().__init__(parent)
        _ensure_i18n()
        self.setWindowTitle(_t("cad.pick.title"))
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        self.resize(1280, 800)
        self.setWindowState(self.windowState() | Qt.WindowMaximized)
        path = Path(dxf_path) if dxf_path else _SAMPLE_DXF
        self.out_dir = out_dir
        self._quiet = False
        self.advance = None
        self.session = session or PickSession.open(str(path))
        self._build_ui()
        self._match_undo_height()
        self._wire()
        self._sync_panel()

    def _match_undo_height(self) -> None:
        """↺ 높이를 옆 버튼에 맞춘다.

        테마 QSS 는 'QDialog QPushButton' 처럼 창 안에 있을 때만 걸리는 규칙을
        쓴다. 버튼을 만들 때는 아직 창에 안 붙어 있어 그때 재면 테마 없는
        높이(26)가 잡히고, 정작 실제 화면에서는 옆 버튼만 24 로 줄어 ↺ 가
        저 혼자 높아진다. 그래서 다 붙은 뒤에 잰다.
        """
        self.btn_cancel.ensurePolished()
        self.btn_undo.setFixedHeight(self.btn_cancel.sizeHint().height())

    def _build_ui(self) -> None:
        self.canvas = self._build_canvas()
        self.panel = self._build_panel()
        body = QHBoxLayout()
        body.addWidget(self.panel)
        body.addWidget(self.canvas, stretch=1)
        self.setLayout(body)

    def _build_panel(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("CadPickPanel")
        panel.setFixedWidth(_PANEL_W)
        vbox = QVBoxLayout(panel)
        vbox.setContentsMargins(0, 0, 0, 0)
        self.lbl_guide = _guide("cad.pick.guide")
        vbox.addWidget(self.lbl_guide)
        vbox.addWidget(self._build_pipe())
        vbox.addWidget(self._build_head())
        vbox.addLayout(self._build_buttons())
        vbox.addStretch(1)
        return panel

    def _build_pipe(self) -> QGroupBox:
        box = QGroupBox(_t("cad.pick.group_pipe"))
        form = QFormLayout()
        self.lbl_desc_pipe = _hint("cad.pick.desc_pipe")
        self.lbl_hint_pipe = _hint("cad.pick.hint_fitting")
        form.addRow(self.lbl_desc_pipe)
        form.addRow(self.lbl_hint_pipe)
        # «배관선택»=찍기 시작 · «선택완료»=재료 완료 → 헤드 해금.
        # 찍기 절차가 그 둘을 다른 일로 다루니 버튼도 둘이다.
        self.row_pipe_btns = QWidget()
        row = QHBoxLayout(self.row_pipe_btns)
        row.setContentsMargins(0, 0, 0, 0)
        self.btn_pipe_select = QPushButton(_t("cad.pick.pipe_select"))
        self.btn_pipe_done = QPushButton(_t("cad.pick.pipe_done"))
        row.addWidget(self.btn_pipe_select, 1)
        row.addWidget(self.btn_pipe_done, 1)
        form.addRow(self.row_pipe_btns)
        box.setLayout(form)
        return box

    def _build_head(self) -> QGroupBox:
        box = QGroupBox(_t("cad.pick.group_head"))
        form = QFormLayout()
        self.lbl_desc_head = _hint("cad.pick.desc_head")
        form.addRow(self.lbl_desc_head)
        slots = QWidget()
        slots.setObjectName("ButtonPanel")
        slot_row = QHBoxLayout(slots)
        slot_row.setContentsMargins(5, 2, 5, 2)
        slot_row.setSpacing(0)
        self.btn_slot_updown = _mode_button(_t("cad.pick.slot_updown"), first=True)
        self.btn_slot_combo = _mode_button(_t("cad.pick.slot_combo"), first=False)
        self.btn_slot_updown.setChecked(True)
        group = QButtonGroup(self)
        group.setExclusive(True)
        group.addButton(self.btn_slot_updown)
        group.addButton(self.btn_slot_combo)
        _bind_checked_contrast(self.btn_slot_updown, self.btn_slot_combo)
        slot_row.addWidget(self.btn_slot_updown)
        slot_row.addWidget(self.btn_slot_combo)
        form.addRow(slots)
        box.setLayout(form)
        box.setEnabled(False)
        self.box_head = box
        return box

    def _build_canvas(self) -> CadPickCanvas:
        return CadPickCanvas(self.session)

    def _build_buttons(self) -> QHBoxLayout:
        row = QHBoxLayout()
        self.btn_undo = QPushButton("↺")
        self.btn_undo.setObjectName("UndoBtn")
        self.btn_undo.setToolTip(_t("cad.pick.undo"))
        self.btn_next = QPushButton(_t("cad.pick.next"))
        self.btn_cancel = QPushButton(_t("취소"))
        self.btn_cancel.clicked.connect(self.reject)
        # UndoBtn 은 글자가 18px 이라 그냥 두면 저 혼자 높아진다.
        # 높이는 창에 다 붙은 뒤 _match_undo_height() 에서 맞춘다.
        self.btn_undo.setFixedWidth(34)
        row.addWidget(self.btn_undo)
        row.addStretch()
        row.addWidget(self.btn_next)
        row.addWidget(self.btn_cancel)
        return row

    def _wire(self) -> None:
        self.btn_pipe_select.clicked.connect(self._on_pipe_select)
        self.btn_pipe_done.clicked.connect(self._on_pipe_done)
        self.btn_slot_updown.clicked.connect(
            lambda: self._on_slot(SLOT_UPDOWN))
        self.btn_slot_combo.clicked.connect(
            lambda: self._on_slot(SLOT_COMBO))
        self.btn_undo.clicked.connect(self._on_undo)
        self.btn_next.clicked.connect(self._on_next)
        self.canvas.picked.connect(self._on_pick)

    def _sync_panel(self) -> None:
        done = self.session.mat_done
        self.box_head.setEnabled(done)
        if done:
            up = self.session.head_label == SLOT_UPDOWN
            self.btn_slot_updown.setChecked(up)
            self.btn_slot_combo.setChecked(not up)

    def _on_pipe_select(self) -> None:
        self.session.select_pipe()
        self._sync_panel()

    def _notice(self, title: str, text: str) -> None:
        if self._quiet:
            return
        QMessageBox.information(self, title, text)

    def _on_pipe_done(self) -> None:
        if not self.session.complete_pipe():
            self._notice(_t("cad.pick.title"), _t("cad.pick.err_no_pipe"))
            return
        self._sync_panel()

    def _on_slot(self, label: str) -> None:
        self.session.set_slot(label)
        self._sync_panel()

    def _on_pick(self, x: float, y: float, max_d: float) -> None:
        got = self.session.click(x, y, max_d=max_d)
        if not got:
            return
        n_clr = got.get("헤드해제") or 0
        if n_clr:
            self._notice(
                _t("cad.pick.heads_cleared_title"),
                _t("cad.pick.heads_cleared").format(n=n_clr))
        self.session.backup(self.out_dir)
        self.canvas.refresh_highlight()
        self._sync_panel()

    def _on_undo(self) -> None:
        if self.session.undo() is None:
            return
        self.session.backup(self.out_dir)
        self.canvas.refresh_highlight()
        self._sync_panel()

    def _on_next(self) -> None:
        try:
            if self.session.commit(self.out_dir) is None:
                self._notice(_t("cad.pick.title"), _t("cad.pick.err_no_pipe"))
                return
        except ValueError as e:
            self._notice(_t("cad.pick.title"), _t(str(e)))
            return
        except Exception as e:
            self._notice(_t("cad.pick.title"), str(e))
            return
        if callable(self.advance):
            self.advance()
            return
        self.accept()


if __name__ == "__main__":
    from ui.theme_manager import apply as apply_theme
    app = QApplication.instance() or QApplication(sys.argv)
    apply_theme(app, "system")
    dlg = CadPickDialog()
    dlg.showMaximized()
    dlg.exec()
