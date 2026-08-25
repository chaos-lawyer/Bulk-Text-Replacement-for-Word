import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import QPoint, QPointF, Qt
    from PySide6.QtGui import QWheelEvent
    from PySide6.QtWidgets import QApplication, QScrollArea, QTableWidget, QVBoxLayout, QWidget
    from ui.widgets.scroll_helper import PreventWheelPropagationFilter, prevent_wheel_propagation

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class ScrollHelperTests(unittest.TestCase):
    def test_prevent_wheel_propagation_absorbs_boundary_wheel_events(self):
        # Create a parent scroll area simulating the page
        parent_scroll = QScrollArea()
        container = QWidget()
        layout = QVBoxLayout(container)

        # Create an inner table widget
        inner_table = QTableWidget(10, 3)
        layout.addWidget(inner_table)
        container.setLayout(layout)
        parent_scroll.setWidget(container)

        # Attach wheel isolation filter
        filt = prevent_wheel_propagation(inner_table)
        self.assertIsInstance(filt, PreventWheelPropagationFilter)

        # Simulate wheel event scrolling UP at the top boundary
        wheel_event_up = QWheelEvent(
            QPointF(10, 10),
            QPointF(10, 10),
            QPoint(0, 0),
            QPoint(0, 120),  # Scroll up
            Qt.NoButton,
            Qt.NoModifier,
            Qt.ScrollUpdate,
            False,
        )

        # Send wheel event to inner table viewport
        handled = filt.eventFilter(inner_table.viewport(), wheel_event_up)
        self.assertTrue(handled)
        self.assertTrue(wheel_event_up.isAccepted())

        # Simulate wheel event scrolling DOWN at the bottom boundary
        wheel_event_down = QWheelEvent(
            QPointF(10, 10),
            QPointF(10, 10),
            QPoint(0, 0),
            QPoint(0, -120),  # Scroll down
            Qt.NoButton,
            Qt.NoModifier,
            Qt.ScrollUpdate,
            False,
        )

        handled_down = filt.eventFilter(inner_table.viewport(), wheel_event_down)
        self.assertTrue(handled_down)
        self.assertTrue(wheel_event_down.isAccepted())

    def test_filter_attached_to_viewport(self):
        table = QTableWidget(5, 5)
        prevent_wheel_propagation(table)
        # Check that event filter is registered
        filters = [child for child in table.children() if isinstance(child, PreventWheelPropagationFilter)]
        self.assertTrue(len(filters) >= 1)
