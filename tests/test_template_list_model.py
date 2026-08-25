import os
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

from core.models import TemplateMergeItem
from ui.models.template_list_model import TemplateListModel

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class TemplateListModelTests(unittest.TestCase):
    def setUp(self):
        self.model = TemplateListModel()

    def test_three_columns_structure(self):
        self.assertEqual(self.model.columnCount(), 3)
        self.assertEqual(self.model.HEADERS, ["启用", "模板名称", "输出路径与文件名规则"])

    def test_add_and_deduplicate_templates_with_default_rule(self):
        p1 = "/path/to/contract_a.docx"
        p2 = "/path/to/contract_b.docx"

        # Add 2 templates
        added = self.model.add_templates([p1, p2])
        self.assertEqual(added, 2)
        self.assertEqual(self.model.rowCount(), 2)

        # Verify default filename rule
        idx_rule = self.model.index(0, TemplateListModel.COL_FILENAME_RULE)
        self.assertEqual(self.model.data(idx_rule, Qt.DisplayRole), "{{模板名}}-{{数据序号}}")

        # Attempting to add duplicate path
        added2 = self.model.add_templates([p1])
        self.assertEqual(added2, 0)
        self.assertEqual(self.model.rowCount(), 2)

    def test_toggle_enabled_and_edit_rule(self):
        self.model.add_templates(["/path/to/my_contract.docx"])
        idx_enabled = self.model.index(0, TemplateListModel.COL_ENABLED)
        idx_name = self.model.index(0, TemplateListModel.COL_NAME)
        idx_rule = self.model.index(0, TemplateListModel.COL_FILENAME_RULE)

        # Initial state
        self.assertEqual(self.model.data(idx_enabled, Qt.CheckStateRole), Qt.Checked)
        self.assertEqual(self.model.data(idx_name, Qt.DisplayRole), "my_contract")

        # Disable
        self.model.setData(idx_enabled, Qt.Unchecked, Qt.CheckStateRole)
        self.assertEqual(len(self.model.get_enabled_templates()), 0)

        # Rename
        self.model.setData(idx_name, "自定义合同", Qt.EditRole)
        self.assertEqual(self.model.data(idx_name, Qt.DisplayRole), "自定义合同")
        self.assertEqual(self.model.get_templates()[0].display_name, "自定义合同")

        # Edit Filename Rule
        self.model.setData(idx_rule, "{{客户名称}}-合同-{{合同编号}}", Qt.EditRole)
        self.assertEqual(self.model.data(idx_rule, Qt.DisplayRole), "{{客户名称}}-合同-{{合同编号}}")
        self.assertEqual(self.model.get_templates()[0].filename_rule, "{{客户名称}}-合同-{{合同编号}}")

    def test_remove_and_clear(self):
        self.model.add_templates(["/path/1.docx", "/path/2.docx", "/path/3.docx"])
        self.assertEqual(self.model.rowCount(), 3)

        # Remove row 1
        self.model.remove_indices([1])
        self.assertEqual(self.model.rowCount(), 2)
        self.assertEqual(self.model.get_templates()[1].file_path, os.path.abspath("/path/3.docx"))

        # Clear
        self.model.clear()
        self.assertEqual(self.model.rowCount(), 0)
