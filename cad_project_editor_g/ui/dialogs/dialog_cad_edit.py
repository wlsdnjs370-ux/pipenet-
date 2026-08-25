# 파일명: ui/dialogs/dialog_cad_edit.py
"""배관/헤드 손질 — 화면만. 판정은 services.cad_import.edit."""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[2]
if __package__ in (None, "") and str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from PySide6.QtCore import QPoint, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import (
    QBrush, QColor, QIcon, QImage, QPainter, QPainterPath, QPen, QPixmap,
)
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QDialog, QFormLayout, QFrame,
    QGraphicsItem, QGraphicsScene, QGraphicsView, QGroupBox, QHBoxLayout,
    QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget,
)

from services.cad_import.colors import (
    C_DARK, EDIT_PENDING, EDIT_SOURCE, EDIT_VALVE, EDIT_WET_PIPE, KIND_COLORS,
)
from services.cad_import.edit import (
    MODE_DELETE, MODE_JOIN, MODE_SOURCE, MODE_VALVE, PICK_PX, EditSession,
)
from services.i18n_service import I18nService, _t
from ui.theme_manager import current_theme as _cur_theme

_PANEL_W = 300
_JOIN_SAMPLE = _ROOT / "docs" / "import" / "_cad_edit_sample_이음.png"
_JOIN_SAMPLE_W = 240
_DRAG_PX = 5.0
# 종류 이름·색은 services.cad_import.colors.KIND_COLORS.
_KIND_BUTTONS = (
    ("상향식", "cad.edit.kind_upright"),
    ("하향식", "cad.edit.kind_pendant"),
    ("상하향식", "cad.edit.kind_combo"),
)
# 미지정으로 «바꾸는» 일은 없다 — 도면 회색이 무엇인지 알리는 글만 둔다.
_KIND_NOTE = ("미지정", "cad.edit.kind_none")
_DOT_PX = 12


def _ensure_i18n() -> None:
    if _t("cad.edit.title") == "cad.edit.title":
        I18nService.instance().load("ko")


def _guide(key: str) -> QLabel:
    """카드 맨 위 사용법 — 단계별 ※ 안내보다 위에 있다는 뜻으로 파랑."""
    lbl = QLabel(_t(key))
    lbl.setWordWrap(True)
    lbl.setObjectName("CadEditGuide")
    lbl.setStyleSheet("QLabel#CadEditGuide { font-weight: bold; }")
    return lbl


def _hint(key: str) -> QLabel:
    """단계 안의 안내 — 회색 처리된 단계에서는 같이 흐려진다."""
    lbl = QLabel(_t(key))
    lbl.setWordWrap(True)
    lbl.setObjectName("CadEditHint")
    lbl.setStyleSheet("QLabel#CadEditHint { font-weight: bold; }")
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


def _mode_button(text: str, *, segpos: str | None = None) -> QPushButton:
    btn = QPushButton(text)
    btn.setCheckable(True)
    if segpos is not None:
        btn.setProperty("segpos", segpos)
    return btn


def _mode_row(*buttons: QPushButton) -> QWidget:
    row = QWidget()
    row.setObjectName("ButtonPanel")
    lay = QHBoxLayout(row)
    lay.setContentsMargins(5, 2, 5, 2)
    lay.setSpacing(0)
    for btn in buttons:
        lay.addWidget(btn)
    return row


def _swatch(color: str, key: str) -> QWidget:
    """종류색 한 칸 — 도면 위 헤드 색과 같은 값이어야 읽힌다."""
    cell = QWidget()
    lay = QHBoxLayout(cell)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(4)
    dot = QLabel()
    dot.setFixedSize(_DOT_PX, _DOT_PX)
    dot.setStyleSheet(
        f"QLabel {{ background: {color}; border-radius: {_DOT_PX // 2}px; }}")
    lay.addWidget(dot)
    lay.addWidget(QLabel(_t(key)))
    lay.addStretch(1)
    return cell


def _join_sample() -> QLabel:
    """이음 예제 그림. 없으면 빈 칸 — 그림 때문에 화면이 죽지는 않는다."""
    lbl = QLabel()
    lbl.setObjectName("CadEditJoinSample")
    lbl.setAlignment(Qt.AlignCenter)
    if not _JOIN_SAMPLE.is_file():
        return lbl
    pix = QPixmap(str(_JOIN_SAMPLE))
    if pix.isNull():
        return lbl
    # 원본은 검정 선+투명. 다크 칸에서는 선만 밝게 뒤집는다.
    if _cur_theme() == "dark":
        img = pix.toImage()
        img.invertPixels(QImage.InvertMode.InvertRgb)
        pix = QPixmap.fromImage(img)
    app = QApplication.instance()
    screen = app.primaryScreen() if app is not None else None
    dpr = float(screen.devicePixelRatio()) if screen is not None else 1.0
    target = max(1, int(_JOIN_SAMPLE_W * dpr))
    if pix.width() != target:
        pix = pix.scaledToWidth(target, Qt.SmoothTransformation)
    pix.setDevicePixelRatio(dpr)
    lbl.setPixmap(pix)
    return lbl


def _dot_icon(color: str) -> QIcon:
    """단추 앞 종류색 점. 확대해도 뭉개지지 않게 크게 그려 아이콘으로 줄인다."""
    pix = QPixmap(4 * _DOT_PX, 4 * _DOT_PX)
    pix.fill(Qt.transparent)
    painter = QPainter(pix)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor(color))
    painter.drawEllipse(pix.rect())
    painter.end()
    return QIcon(pix)


class CadEditCanvas(QGraphicsView):
    """다크 망. 클릭=손질 · 끌기=이동 · 휠=확대. 판정은 세션."""

    picked = Signal(float, float, float)

    def __init__(self, session: EditSession, parent=None):
        super().__init__(parent)
        self.setObjectName("CadEditCanvas")
        self.setFrameShape(QFrame.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setBackgroundBrush(QColor(16, 16, 16))
        self.setScene(QGraphicsScene(self))
        # 카메라는 도면 items 박스에 가두지 않는다. 맞춤은 itemsBoundingRect.
        self.setSceneRect(QRectF(-1e9, -1e9, 2e9, 2e9))
        self._session = session
        self._user_view = False
        self._drag = {"on": False, "moved": False, "pos": QPoint()}
        self._wet_items = []
        self._head_items = []
        self.redraw()

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        if not self._user_view:
            self._fit()

    def wheelEvent(self, ev):
        if ev.angleDelta().y() == 0:
            return
        self._user_view = True
        k = 1.25 if ev.angleDelta().y() > 0 else 0.8
        # 찍기 AnchorUnderMouse 는 거대 sceneRect 에서 첫 휠에 원점(0,0)으로 붙는다.
        # 찍기 팬과 같은 mapToScene 차이로 마우스 아래 점을 유지한다.
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

    def redraw(self) -> None:
        scene = self.scene()
        self._wet_items = []
        self._head_items = []
        scene.clear()
        scene.setSceneRect(QRectF(-1e9, -1e9, 2e9, 2e9))
        geom = self._session.display_geom()
        for segs, color in geom["body_groups"]:
            pen = QPen(QColor(color))
            pen.setCosmetic(True)
            path = QPainterPath()
            for a, b in segs:
                path.moveTo(a[0], -a[1])
                path.lineTo(b[0], -b[1])
            scene.addPath(path, pen)
        self._paint_wet_overlay(geom.get("wet_pipes") or ())
        pending = geom.get("pending")
        if pending:
            pen = QPen(QColor(EDIT_PENDING))
            pen.setCosmetic(True)
            pen.setWidth(2)
            a, b = pending
            scene.addLine(a[0], -a[1], b[0], -b[1], pen)
        for disk, color in geom["heads"]:
            x, y, r = disk[0], disk[1], disk[2]
            it = scene.addEllipse(x - r, -y - r, 2 * r, 2 * r,
                                 QPen(Qt.NoPen), QBrush(QColor(color)))
            self._head_items.append(it)
        # 팔 후보가 여럿인 헤드 — 프로그램이 중심 최근접 끝으로 이은 자리.
        # 고른 헤드(selected_head)와 같은 테두리 패턴, 색만 «안붙은끝점».
        for disk in geom.get("multi_heads") or ():
            x, y, r = disk[0], disk[1], disk[2]
            pen = QPen(QColor(C_DARK["안붙은끝점"]))
            pen.setCosmetic(True)
            pen.setWidth(2)
            scene.addEllipse(x - r, -y - r, 2 * r, 2 * r, pen)
        sel = geom.get("selected_head")
        if sel:
            x, y, r = sel[0], sel[1], sel[2]
            pen = QPen(QColor(C_DARK["헤드원"]))
            pen.setCosmetic(True)
            pen.setWidth(2)
            scene.addEllipse(x - r, -y - r, 2 * r, 2 * r, pen)
        for x, y in geom["sources"]:
            it = scene.addEllipse(-5, -5, 10, 10, QPen(Qt.NoPen),
                                 QBrush(QColor(EDIT_SOURCE)))
            it.setPos(x, -y)
            it.setFlag(QGraphicsItem.ItemIgnoresTransformations)
        for x, y in geom["valves"]:
            it = scene.addRect(-5, -5, 10, 10, QPen(Qt.NoPen),
                               QBrush(QColor(EDIT_VALVE)))
            it.setPos(x, -y)
            it.setFlag(QGraphicsItem.ItemIgnoresTransformations)

    def _paint_wet_overlay(self, segs) -> None:
        """젖은 배관은 독립 선으로 갱신해 큰 경로의 전면 재도장을 피한다.

        하나의 계속 커지는 QGraphicsPathItem은 setPath 때마다 그 경계 전체가
        dirty 영역이 된다. 큰 도면에서는 아래 배관망이 다시 칠해지는 순간이
        흰 선 깜빡임처럼 보이므로, 연출 오버레이만 작은 선 객체를 유지한다.
        """
        scene = self.scene()
        wet_col = QColor(EDIT_WET_PIPE)
        wet_col.setAlphaF(0.45)
        wet_pen = QPen(wet_col)
        wet_pen.setCosmetic(True)
        wet_pen.setWidthF(1.4)
        n = len(segs)
        while len(self._wet_items) < n:
            self._wet_items.append(scene.addLine(0, 0, 0, 0, wet_pen))
        for i, (a, b) in enumerate(segs):
            item = self._wet_items[i]
            target = (a[0], -a[1], b[0], -b[1])
            line = item.line()
            if (line.x1(), line.y1(), line.x2(), line.y2()) != target:
                item.setLine(*target)
            if not item.isVisible():
                item.setVisible(True)
        for item in self._wet_items[n:]:
            item.setVisible(False)

    def paint_flow_frame(self) -> None:
        """연출 틱 — 망은 두고 초록 관·헤드 색만."""
        geom = self._session.display_geom(net=False)
        self._paint_wet_overlay(geom.get("wet_pipes") or ())
        for it, (_disk, color) in zip(self._head_items, geom["heads"]):
            it.setBrush(QBrush(QColor(color)))


class CadEditDialog(QDialog):
    def __init__(self, parent=None, key="3F", out_dir=None, session=None):
        super().__init__(parent)
        _ensure_i18n()
        self.setWindowTitle(_t("cad.edit.title"))
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
        self.resize(1280, 800)
        self.setWindowState(self.windowState() | Qt.WindowMaximized)
        self.out_dir = out_dir
        self._quiet = False
        self.went_back = False
        self.advance = None
        self.session = session or EditSession.open(key, out_dir=out_dir)
        self._build_ui()
        self._match_undo_height()
        self._wire()
        self._flow_timer = QTimer(self)
        self._flow_timer.setInterval(32)
        self._flow_timer.timeout.connect(self._on_flow_tick)
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
        panel.setObjectName("CadEditPanel")
        panel.setFixedWidth(_PANEL_W)
        vbox = QVBoxLayout(panel)
        vbox.setContentsMargins(0, 0, 0, 0)
        # 모드는 하나만 켜진다 — 이음·삭제·급수시작위치·알람밸브위치
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)
        self.lbl_guide = _guide("cad.edit.guide")
        vbox.addWidget(self.lbl_guide)
        vbox.addWidget(self._build_pipe())
        vbox.addWidget(self._build_head())
        vbox.addWidget(self._build_water())
        self.lbl_note_source = _hint("cad.edit.note_source")
        vbox.addWidget(self.lbl_note_source)
        vbox.addLayout(self._build_buttons())
        vbox.addStretch(1)
        _bind_checked_contrast(
            self.btn_mode_join, self.btn_mode_delete,
            self.btn_mode_source, self.btn_mode_valve)
        return panel

    def _build_pipe(self) -> QGroupBox:
        box = QGroupBox(_t("cad.edit.group_pipe"))
        form = QFormLayout()
        self.lbl_desc_pipe = _hint("cad.edit.desc_pipe")
        form.addRow(self.lbl_desc_pipe)
        self.lbl_join_sample = _join_sample()
        form.addRow(self.lbl_join_sample)
        self.btn_mode_join = _mode_button(_t("cad.edit.mode_join"),
                                          segpos="first")
        self.btn_mode_delete = _mode_button(_t("cad.edit.mode_delete"),
                                            segpos="last")
        self.btn_mode_join.setChecked(True)
        self.mode_group.addButton(self.btn_mode_join)
        self.mode_group.addButton(self.btn_mode_delete)
        self.row_pipe_modes = _mode_row(self.btn_mode_join,
                                        self.btn_mode_delete)
        form.addRow(self.row_pipe_modes)
        box.setLayout(form)
        return box

    def _build_head(self) -> QGroupBox:
        box = QGroupBox(_t("cad.edit.group_head"))
        form = QFormLayout()
        self.lbl_desc_head = _hint("cad.edit.desc_head")
        form.addRow(self.lbl_desc_head)
        # 고른 헤드를 그 종류로 «바꾸는» 단추다. 모드가 아니므로 눌린 채 남지 않는다.
        self.btn_kind = {}
        self.lbl_kind_count = {}
        for kind, key in _KIND_BUTTONS:
            btn = QPushButton(_t(key))
            btn.setIcon(_dot_icon(KIND_COLORS[kind]))
            btn.setIconSize(QSize(_DOT_PX, _DOT_PX))
            # 물흐름 테스트가 젖다고 판정한 헤드만 센다 — 전체 개수가 아니다.
            count = QLabel("0")
            count.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            count.setMinimumWidth(34)
            row = QWidget()
            lay = QHBoxLayout(row)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.addWidget(btn, 1)
            lay.addWidget(count)
            form.addRow(row)
            self.btn_kind[kind] = btn
            self.lbl_kind_count[kind] = count
        kind, key = _KIND_NOTE
        self.lbl_kind_none = _swatch(KIND_COLORS[kind], key)
        form.addRow(self.lbl_kind_none)
        box.setLayout(form)
        self.box_head = box
        return box

    def _build_water(self) -> QGroupBox:
        box = QGroupBox(_t("cad.edit.group_water"))
        form = QFormLayout()
        self.lbl_desc_water = _hint("cad.edit.desc_water")
        form.addRow(self.lbl_desc_water)
        self.btn_mode_source = _mode_button(_t("cad.edit.mode_source"),
                                            segpos="first")
        self.btn_mode_valve = _mode_button(_t("cad.edit.mode_valve"),
                                           segpos="last")
        self.mode_group.addButton(self.btn_mode_source)
        self.mode_group.addButton(self.btn_mode_valve)
        self.row_water_modes = _mode_row(self.btn_mode_source,
                                         self.btn_mode_valve)
        form.addRow(self.row_water_modes)
        # 찍기와 물흐름은 다른 일이다 — 모드가 아니라 한 번 도는 단추.
        self.btn_flow = QPushButton(_t("cad.edit.flow"))
        form.addRow(self.btn_flow)
        box.setLayout(form)
        self.box_water = box
        return box

    def _build_canvas(self) -> CadEditCanvas:
        return CadEditCanvas(self.session)

    def _build_buttons(self) -> QHBoxLayout:
        row = QHBoxLayout()
        self.btn_undo = QPushButton("↺")
        self.btn_undo.setObjectName("UndoBtn")
        self.btn_undo.setToolTip(_t("cad.edit.undo"))
        self.btn_back = QPushButton(_t("cad.edit.back"))
        self.btn_next = QPushButton(_t("cad.edit.next"))
        self.btn_cancel = QPushButton(_t("취소"))
        self.btn_cancel.clicked.connect(self.reject)
        # UndoBtn 은 글자가 18px 이라 그냥 두면 저 혼자 높아진다.
        # 높이는 창에 다 붙은 뒤 _match_undo_height() 에서 맞춘다.
        self.btn_undo.setFixedWidth(34)
        row.addWidget(self.btn_undo)
        row.addStretch()
        row.addWidget(self.btn_back)
        row.addWidget(self.btn_next)
        row.addWidget(self.btn_cancel)
        return row

    def _wire(self) -> None:
        self.btn_mode_join.clicked.connect(lambda: self._on_mode(MODE_JOIN))
        self.btn_mode_delete.clicked.connect(lambda: self._on_mode(MODE_DELETE))
        self.btn_mode_source.clicked.connect(lambda: self._on_mode(MODE_SOURCE))
        self.btn_mode_valve.clicked.connect(lambda: self._on_mode(MODE_VALVE))
        for kind, btn in self.btn_kind.items():
            btn.clicked.connect(lambda _=False, k=kind: self._on_kind(k))
        self.btn_flow.clicked.connect(self._on_flow)
        self.btn_undo.clicked.connect(self._on_undo)
        self.btn_back.clicked.connect(self._on_back)
        self.btn_next.clicked.connect(self._on_next)
        self.canvas.picked.connect(self._on_pick)

    def _sync_panel(self) -> None:
        counts = self.session.wet_kind_counts()
        for kind, lbl in self.lbl_kind_count.items():
            lbl.setText(str(counts.get(kind, 0)))

    def _refresh(self) -> None:
        self.canvas.redraw()
        self._sync_panel()

    def _notice(self, title: str, text: str) -> None:
        if self._quiet:
            return
        QMessageBox.information(self, title, text)

    def _on_mode(self, mode: str) -> None:
        self.session.set_mode(mode)
        self._refresh()

    def _on_pick(self, x: float, y: float, max_d: float) -> None:
        source_shot = self.session.mode == MODE_SOURCE
        if self.session.click(x, y, max_d) is None:
            return
        if source_shot:
            self.session.set_mode(MODE_JOIN)
            self.btn_mode_join.setChecked(True)
        self._refresh()

    def _on_kind(self, kind: str) -> None:
        if self.session.set_kind(kind) is None:
            return
        self._refresh()

    def _on_flow(self) -> None:
        self.session.flow()
        self._refresh()
        if self.session.flow_animating():
            self._flow_timer.start()
        else:
            self._flow_timer.stop()

    def _on_flow_tick(self) -> None:
        more = self.session.flow_tick()
        if more:
            self.canvas.paint_flow_frame()
            return
        self._flow_timer.stop()
        self._refresh()

    def _on_undo(self) -> None:
        if not self.session.undo():
            return
        self._refresh()

    def _on_back(self) -> None:
        self.went_back = True
        self.reject()

    def _on_next(self) -> None:
        try:
            self.session.commit(self.out_dir)
        except ValueError as e:
            self._notice(_t("cad.edit.title"), _t(str(e)))
            return
        except Exception as e:
            self._notice(_t("cad.edit.title"), str(e))
            return
        if callable(self.advance):
            self.advance()
            return
        self.accept()


if __name__ == "__main__":
    from ui.theme_manager import apply as apply_theme
    app = QApplication.instance() or QApplication(sys.argv)
    apply_theme(app, "system")
    dlg = CadEditDialog()
    dlg.showMaximized()
    dlg.exec()
