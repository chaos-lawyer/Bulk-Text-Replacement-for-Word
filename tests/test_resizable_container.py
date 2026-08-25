import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QTableView
    from ui.widgets.resizable_container import ResizableTableContainer, TableResizeGrip

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class ResizableContainerTests(unittest.TestCase):
    def test_resizable_container_init_and_resize(self):
        table = QTableView()
        container = ResizableTableContainer(
            target_widget=table,
            initial_height=120,
            min_height=80,
            max_height=400,
        )

        self.assertEqual(container.get_table_height(), 120)

        # Height clamp within bounds
        container.set_table_height(250)
        self.assertEqual(container.get_table_height(), 250)

        # Clamp min
        container.set_table_height(40)
        self.assertEqual(container.get_table_height(), 80)

        # Clamp max
        container.set_table_height(600)
        self.assertEqual(container.get_table_height(), 400)

        self.assertIsInstance(container.grip, TableResizeGrip)
