"""Helper utilities for isolated scroll behavior on embedded scrollable widgets."""

from __future__ import annotations

try:
    from PySide6.QtCore import QEvent, QObject
    from PySide6.QtWidgets import QAbstractScrollArea

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class PreventWheelPropagationFilter(QObject):
        """Event filter that intercepts wheel events on a scrollable view or its viewport,

        preventing the event from bubbling up to parent scroll containers when reaching top/bottom bounds.
        """

        def __init__(self, scroll_area: QAbstractScrollArea):
            super().__init__(scroll_area)
            self._scroll_area = scroll_area

        def eventFilter(self, obj: QObject, event: QEvent) -> bool:
            if event.type() == QEvent.Wheel:
                # Forward to scroll area to handle internal scrollbar offset
                self._scroll_area.wheelEvent(event)
                # Mark accepted and consume so Qt will never propagate to parent scroll area
                event.accept()
                return True
            return False


    def prevent_wheel_propagation(scroll_area: QAbstractScrollArea) -> PreventWheelPropagationFilter:
        """Attach wheel isolation filter to a scroll area and its viewport."""
        filter_obj = PreventWheelPropagationFilter(scroll_area)
        scroll_area.installEventFilter(filter_obj)
        if hasattr(scroll_area, "viewport") and scroll_area.viewport() is not None:
            scroll_area.viewport().installEventFilter(filter_obj)
        return filter_obj

else:

    class PreventWheelPropagationFilter:  # type: ignore
        pass

    def prevent_wheel_propagation(scroll_area):  # type: ignore
        return None
