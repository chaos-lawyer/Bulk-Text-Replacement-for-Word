"""Resizable container for tables allowing users to drag the bottom-right handle to adjust table height."""

from __future__ import annotations

from ui.theme.theme_manager import THEME

try:
    from PySide6.QtCore import Qt, Signal
    from PySide6.QtGui import QColor, QMouseEvent, QPainter, QPaintEvent, QPen
    from PySide6.QtWidgets import (
        QAbstractScrollArea,
        QHBoxLayout,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )
    from ui.widgets.scroll_helper import prevent_wheel_propagation

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class TableResizeGrip(QWidget):
        """Corner resize grip widget positioned at the bottom-right corner of a table."""

        def __init__(self, parent: ResizableTableContainer | None = None):
            super().__init__(parent)
            self._container = parent
            self._is_dragging = False
            self._drag_start_pos_y = 0.0
            self._drag_start_height = 0

            self.setFixedSize(24, 14)
            self.setCursor(Qt.SizeVerCursor)
            self.setToolTip("按住并上下拖动以调整表格高度")

        def mousePressEvent(self, event: QMouseEvent):
            if event.button() == Qt.LeftButton and self._container:
                self._is_dragging = True
                self._drag_start_pos_y = event.globalPosition().y()
                self._drag_start_height = self._container.target_widget.height()
                event.accept()
                return
            super().mousePressEvent(event)

        def mouseMoveEvent(self, event: QMouseEvent):
            if self._is_dragging and self._container:
                delta_y = event.globalPosition().y() - self._drag_start_pos_y
                new_h = int(self._drag_start_height + delta_y)
                self._container.set_table_height(new_h)
                event.accept()
                return
            super().mouseMoveEvent(event)

        def mouseReleaseEvent(self, event: QMouseEvent):
            if self._is_dragging:
                self._is_dragging = False
                event.accept()
                return
            super().mouseReleaseEvent(event)

        def paintEvent(self, event: QPaintEvent):
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing)

            color = QColor(THEME.tokens.border_strong if not self._is_dragging else THEME.tokens.accent)
            pen = QPen(color, 1.5)
            pen.setCapStyle(Qt.RoundCap)
            painter.setPen(pen)

            w, h = self.width(), self.height()
            # Draw subtle diagonal gripper dots / lines at bottom right
            # Line 1 (shortest)
            painter.drawLine(w - 5, h - 3, w - 3, h - 5)
            # Line 2 (medium)
            painter.drawLine(w - 9, h - 3, w - 3, h - 9)
            # Line 3 (longest)
            painter.drawLine(w - 13, h - 3, w - 3, h - 13)

            painter.end()


    class ResizableTableContainer(QWidget):
        """Container wrapping a table widget with a draggable resize grip at the bottom-right."""

        heightChanged = Signal(int)

        def __init__(
            self,
            target_widget: QWidget,
            initial_height: int = 120,
            min_height: int = 70,
            max_height: int = 650,
            parent: QWidget | None = None,
        ):
            super().__init__(parent)
            self.target_widget = target_widget
            self.min_height = min_height
            self.max_height = max_height

            self._build_ui(initial_height)

        def _build_ui(self, initial_height: int):
            layout = QVBoxLayout(self)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(1)

            # Target table
            self.target_widget.setFixedHeight(initial_height)
            self.target_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            if isinstance(self.target_widget, QAbstractScrollArea):
                prevent_wheel_propagation(self.target_widget)
            layout.addWidget(self.target_widget)

            # Bottom bar containing the grip at right
            grip_bar = QHBoxLayout()
            grip_bar.setContentsMargins(0, 0, 2, 0)
            grip_bar.setSpacing(0)
            grip_bar.addStretch()

            self.grip = TableResizeGrip(parent=self)
            grip_bar.addWidget(self.grip)

            layout.addLayout(grip_bar)

        def set_table_height(self, height: int):
            clamped = max(self.min_height, min(self.max_height, height))
            self.target_widget.setFixedHeight(clamped)
            self.heightChanged.emit(clamped)

        def get_table_height(self) -> int:
            return self.target_widget.height()

else:

    class TableResizeGrip:  # type: ignore
        pass

    class ResizableTableContainer:  # type: ignore
        pass
