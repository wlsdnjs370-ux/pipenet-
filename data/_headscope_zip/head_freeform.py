"""헤드 자유형(올가미) 범위 선택 — 다른 프로그램에 붙이는 복사본.

K-Fire 본체(view_graphics.py)의 마우스 범위 선택만 옮겼다.
본체 파일은 수정하지 않았다. 이 묶음을 본체에 넣지 마라.

동작
    1. 모드 ON → 커서 십자
    2. 왼쪽 드래그로 자유형 경로를 그린다 (주황 점선)
    3. 버튼을 놓으면 경로를 닫힌 다각형으로 본다
    4. 중심점이 다각형 안인 대상 id 목록을 콜백으로 돌려준다
    5. ESC 또는 드래그 중 우클릭 → 취소

넣지 않은 것 (K-Fire 전용)
    다중선택 저장소, 우클릭 메뉴, 일괄 속성 변경,
    배관 자유형, 물도달시간 헤드 클릭, 번역 문자열.

붙이는 쪽은 get_heads() 로 「고를 점」만 넘기면 된다.
헤드만 넘기면 헤드만 골라진다. 노즐을 넣으면 노즐도 골라진다.
"""
from __future__ import annotations

from collections.abc import Callable, Iterable
from typing import Sequence

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPainterPath, QPen, QPolygonF
from PySide6.QtWidgets import QGraphicsPathItem, QGraphicsView

HeadPoint = tuple[str, QPointF]


def ids_inside_polygon(
    points: Sequence[QPointF],
    heads: Iterable[HeadPoint],
) -> list[str]:
    """닫힌 다각형 안에 중심점이 들어온 id 목록.

    판정은 Qt OddEvenFill (짝홀 규칙). 자기 자신과 겹친 올가미에서도
    K-Fire와 같은 결과가 나온다. UI는 건드리지 않는다.
    """
    if len(points) < 3:
        return []
    polygon = QPolygonF(list(points))
    result: list[str] = []
    for name, center in heads:
        if polygon.containsPoint(center, Qt.OddEvenFill):
            result.append(name)
    return result


class HeadFreeformSelector:
    """QGraphicsView에 붙이는 자유형 선택 도구.

    view의 mousePress/Move/Release/DoubleClick/keyPress 에서
    handle_* 가 True 를 반환하면 그 이벤트를 먹지 말고 return 하면 된다.
    """

    def __init__(
        self,
        view: QGraphicsView,
        *,
        get_heads: Callable[[], Iterable[HeadPoint]],
        on_selected: Callable[[list[str]], None] | None = None,
        on_empty: Callable[[], None] | None = None,
        on_mode_changed: Callable[[bool], None] | None = None,
    ) -> None:
        self.view = view
        self.get_heads = get_heads
        self.on_selected = on_selected
        self.on_empty = on_empty
        self.on_mode_changed = on_mode_changed

        self.mode = False
        self._active = False
        self._points: list[QPointF] = []
        self._preview: QGraphicsPathItem | None = None

    # ── 모드 ──────────────────────────────────────────────

    def set_mode(self, enabled: bool) -> None:
        if not enabled:
            self.cancel_drag()
        self.mode = bool(enabled)
        self.view.setCursor(Qt.CrossCursor if self.mode else Qt.ArrowCursor)
        if self.mode:
            # 버튼에 포커스가 남으면 ESC가 view로 안 온다
            self.view.setFocus(Qt.OtherFocusReason)
        if self.on_mode_changed is not None:
            self.on_mode_changed(self.mode)

    def cancel_mode(self) -> None:
        """ESC — 드래그·미리보기 지우고 모드 OFF."""
        self.cancel_drag()
        if not self.mode:
            return
        self.mode = False
        self.view.setCursor(Qt.ArrowCursor)
        if self.on_mode_changed is not None:
            self.on_mode_changed(False)

    def cancel_drag(self) -> None:
        """현재 드래그만 취소. 모드는 유지."""
        self._active = False
        self._points = []
        self._remove_preview()

    # ── 이벤트 (True = 이 도구가 먹었다) ─────────────────

    def handle_press(self, event) -> bool:
        if not self.mode:
            return False
        if event.button() == Qt.LeftButton:
            self._begin(self._scene_pos(event))
            event.accept()
            return True
        if event.button() == Qt.RightButton and self._active:
            # 드래그 중 우클릭 → 이번 드래그만 취소, 모드 유지
            self.cancel_drag()
            event.accept()
            return True
        return False

    def handle_move(self, event) -> bool:
        if not self._active:
            return False
        self._append(self._scene_pos(event))
        event.accept()
        return True

    def handle_release(self, event) -> bool:
        if not (self._active and event.button() == Qt.LeftButton):
            return False
        self._finish()
        event.accept()
        return True

    def handle_double_click(self, event) -> bool:
        if not self.mode:
            return False
        event.accept()
        return True

    def handle_key(self, event) -> bool:
        if event.key() != Qt.Key_Escape:
            return False
        if not (self.mode or self._preview is not None):
            return False
        self.cancel_mode()
        event.accept()
        return True

    # ── 내부 (K-Fire와 같은 순서) ─────────────────────────

    def _scene_pos(self, event) -> QPointF:
        return self.view.mapToScene(event.position().toPoint())

    def _begin(self, scene_pos: QPointF) -> None:
        self._active = True
        self._points = [scene_pos]
        self._remove_preview()
        preview = QGraphicsPathItem()
        pen = QPen(QColor(255, 120, 0))
        pen.setStyle(Qt.DashLine)
        pen.setWidthF(1.5)
        pen.setCosmetic(True)
        preview.setPen(pen)
        preview.setZValue(99999)
        scene = self.view.scene()
        if scene is not None:
            scene.addItem(preview)
            self._preview = preview

    def _append(self, scene_pos: QPointF) -> None:
        # 드래그 중에는 점만 쌓고, 포함 판정은 하지 않는다
        self._points.append(scene_pos)
        self._update_preview()

    def _update_preview(self) -> None:
        if self._preview is None or not self._points:
            return
        path = QPainterPath()
        path.moveTo(self._points[0])
        for pt in self._points[1:]:
            path.lineTo(pt)
        try:
            self._preview.setPath(path)
        except RuntimeError:
            self._preview = None

    def _finish(self) -> None:
        self._active = False
        points = list(self._points)
        self._points = []

        if len(points) < 3:
            self._remove_preview()
            return

        closed = QPainterPath()
        closed.moveTo(points[0])
        for pt in points[1:]:
            closed.lineTo(pt)
        closed.closeSubpath()
        if self._preview is not None:
            try:
                self._preview.setPath(closed)
            except RuntimeError:
                self._preview = None

        names = ids_inside_polygon(points, self.get_heads())
        if names:
            # 본체는 여기서 우클릭 메뉴를 띄운 뒤 outline을 지운다.
            # 다른 프로그램에는 메뉴가 없으므로 결과만 넘기고 outline을 치운다.
            self._remove_preview()
            self.mode = False
            self.view.setCursor(Qt.ArrowCursor)
            if self.on_mode_changed is not None:
                self.on_mode_changed(False)
            if self.on_selected is not None:
                self.on_selected(names)
            return

        self._remove_preview()
        if self.on_empty is not None:
            self.on_empty()

    def _remove_preview(self) -> None:
        item = self._preview
        self._preview = None
        if item is None:
            return
        try:
            item_scene = item.scene()
        except RuntimeError:
            return
        if item_scene is None:
            return
        try:
            item_scene.removeItem(item)
        except RuntimeError:
            return
