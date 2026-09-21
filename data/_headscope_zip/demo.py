"""이 묶음만으로 동작을 확인하는 작은 창.

    python demo.py

원 20개가 헤드다. [헤드선택]을 누른 뒤 마우스로 범위를 그리면
안에 들어온 원만 주황으로 바뀐다. ESC로 모드를 끈다.
"""
from __future__ import annotations

import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication,
    QGraphicsEllipseItem,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from head_freeform import HeadFreeformSelector

_IDLE = QColor(70, 140, 220)
_PICKED = QColor(255, 120, 0)


class DemoView(QGraphicsView):
    def __init__(self, selector_holder: dict) -> None:
        super().__init__()
        self._holder = selector_holder
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.Antialiasing)
        self.setBackgroundBrush(QColor(28, 30, 36))
        self.setMouseTracking(True)

    @property
    def lasso(self) -> HeadFreeformSelector:
        return self._holder["lasso"]

    def mousePressEvent(self, event):
        if self.lasso.handle_press(event):
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.lasso.handle_move(event):
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.lasso.handle_release(event):
            return
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if self.lasso.handle_double_click(event):
            return
        super().mouseDoubleClickEvent(event)

    def keyPressEvent(self, event):
        if self.lasso.handle_key(event):
            return
        super().keyPressEvent(event)


def _make_heads(scene: QGraphicsScene) -> dict[str, QGraphicsEllipseItem]:
    items: dict[str, QGraphicsEllipseItem] = {}
    r = 10
    # 4 x 5 격자
    for row in range(4):
        for col in range(5):
            name = f"H{row * 5 + col + 1}"
            x = 80 + col * 90
            y = 70 + row * 80
            item = QGraphicsEllipseItem(-r, -r, r * 2, r * 2)
            item.setPos(x, y)
            item.setBrush(QBrush(_IDLE))
            item.setPen(QPen(QColor(20, 20, 20)))
            item.setToolTip(name)
            scene.addItem(item)
            items[name] = item
    return items


def main() -> None:
    app = QApplication(sys.argv)

    holder: dict = {}
    view = DemoView(holder)
    heads = _make_heads(view.scene())
    view.setSceneRect(0, 0, 560, 400)

    status = QLabel("헤드선택을 누른 뒤, 마우스로 범위를 그리세요.")
    btn = QPushButton("헤드선택")
    btn.setCheckable(True)

    def paint(selected: list[str]) -> None:
        picked = set(selected)
        for name, item in heads.items():
            item.setBrush(QBrush(_PICKED if name in picked else _IDLE))
        if selected:
            status.setText(f"{len(selected)}개 선택: {', '.join(selected)}")
        else:
            status.setText("선택 범위에 헤드가 없습니다. 모드는 그대로입니다.")

    def get_heads():
        return [
            (name, item.sceneBoundingRect().center())
            for name, item in heads.items()
        ]

    lasso = HeadFreeformSelector(
        view,
        get_heads=get_heads,
        on_selected=paint,
        on_empty=lambda: paint([]),
        on_mode_changed=btn.setChecked,
    )
    holder["lasso"] = lasso

    btn.clicked.connect(lasso.set_mode)

    win = QWidget()
    win.setWindowTitle("헤드 범위 선택 — 추출 데모")
    root = QVBoxLayout(win)
    bar = QHBoxLayout()
    bar.addWidget(btn)
    bar.addWidget(status, 1)
    root.addLayout(bar)
    root.addWidget(view)
    win.resize(620, 480)
    win.show()
    view.setFocus(Qt.OtherFocusReason)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
