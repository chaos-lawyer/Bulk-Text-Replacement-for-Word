import os
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
    from ui.widgets.capsule_edit import DropLineEdit, FileCapsuleEdit, FileCapsuleItem

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class CapsuleEditTests(unittest.TestCase):
    def test_file_capsule_edit_set_and_remove(self):
        edit = FileCapsuleEdit(placeholder_text="拖放模板到此处...")
        self.assertFalse(edit.lbl_placeholder.isHidden())
        self.assertTrue(edit.capsule_container.isHidden())

        # Set files
        files = ["/path/to/contract_a.docx", "/path/to/agreement_b.docx"]
        edit.set_files(files)

        self.assertTrue(edit.lbl_placeholder.isHidden())
        self.assertFalse(edit.capsule_container.isHidden())
        self.assertEqual(len(edit._file_paths), 2)

        # Test remove signal emission
        removed_paths = []
        edit.fileRemoved.connect(removed_paths.append)

        # Find first capsule and trigger remove
        first_capsule = edit.capsule_layout.itemAt(0).widget()
        self.assertIsInstance(first_capsule, FileCapsuleItem)
        first_capsule.btn_close.click()

        self.assertEqual(len(removed_paths), 1)
        self.assertEqual(removed_paths[0], "/path/to/contract_a.docx")

    def test_drop_line_edit_creation(self):
        drop_edit = DropLineEdit(supported_extensions=(".xlsx", ".csv"))
        self.assertTrue(drop_edit.acceptDrops())
        drop_edit.setText("/path/to/data.xlsx")
        self.assertEqual(drop_edit.text(), "/path/to/data.xlsx")
