"""클릭(포커스 진입) 시 숫자를 전체선택하는 QDoubleSpinBox."""
from PySide6.QtWidgets import QDoubleSpinBox
from PySide6.QtCore import QTimer


class SelectAllDoubleSpinBox(QDoubleSpinBox):
    """클릭(포커스 진입) 시 숫자를 전체선택하여 즉시 입력 가능하게 한다."""
    def focusInEvent(self, event):
        super().focusInEvent(event)
        QTimer.singleShot(0, self.selectAll)
