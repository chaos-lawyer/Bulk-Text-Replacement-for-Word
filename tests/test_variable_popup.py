import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QLineEdit, QSizePolicy
    from ui.widgets.variable_popup import VariableCapsuleButton, VariableInsertPanel, create_variable_icon

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class VariablePanelTests(unittest.TestCase):
    def test_variable_panel_catalog_and_selection_signal(self):
        panel = VariableInsertPanel()
        panel.set_target_name("采购合同")
        self.assertIn("采购合同", panel.lbl_target.text())

        sys_vars = ["{{模板名}}", "{{数据序号}}"]
        table_vars = ["客户姓名", "合同金额"]

        panel.set_catalog(sys_vars, table_vars)

        # Test variable selection signal
        received_tags = []
        panel.variableSelected.connect(received_tags.append)

        # Trigger capsule insertion
        panel._on_capsule_clicked("{{客户姓名}}")
        self.assertEqual(received_tags, ["{{客户姓名}}"])

        panel._on_capsule_clicked("{{数据序号}}")
        self.assertEqual(received_tags, ["{{客户姓名}}", "{{数据序号}}"])

    def test_variable_panel_search_filtering(self):
        panel = VariableInsertPanel()
        panel.show()
        sys_vars = ["{{模板名}}", "{{数据序号}}"]
        table_vars = ["客户姓名", "合同金额", "地区代码"]
        panel.set_catalog(sys_vars, table_vars)

        # Search for "金额"
        panel.edt_search.setText("金额")
        visible_tbl = [b.tag for b in panel._tbl_buttons if not b.isHidden()]
        self.assertEqual(visible_tbl, ["{{合同金额}}"])
        self.assertFalse(panel.lbl_empty_search.isVisible())

        # Search for non-existent variable
        panel.edt_search.setText("xyz_not_found")
        self.assertTrue(panel.lbl_empty_search.isVisible())

        # Reset search
        panel.reset_search()
        visible_tbl_all = [b.tag for b in panel._tbl_buttons if not b.isHidden()]
        self.assertEqual(len(visible_tbl_all), 3)

    def test_variable_panel_large_scale_bounded_height(self):
        """Verify a large catalog scrolls without exceeding MAX_HEIGHT."""
        panel = VariableInsertPanel()
        sys_vars = ["{{模板名}}"]
        table_vars = [f"列_{i}" for i in range(500)]

        panel.set_catalog(sys_vars, table_vars)

        self.assertGreaterEqual(panel.height(), panel.MIN_HEIGHT)
        self.assertLessEqual(panel.height(), panel.MAX_HEIGHT)

    def test_variable_panel_auto_scaling_height(self):
        """Verify panel height expands with more content and shrinks with less content."""
        panel = VariableInsertPanel()
        panel.resize(400, 100)

        # Small content (only 2 system variables)
        panel.set_catalog(["{{模板名}}", "{{数据序号}}"], [])
        small_height = panel.height()
        self.assertGreaterEqual(small_height, panel.MIN_HEIGHT)
        self.assertLessEqual(small_height, 130)

        # Larger content (20 variables across multiple rows)
        panel.set_catalog(
            ["{{模板名}}", "{{数据序号}}", "{{日期}}"],
            [f"字段_{i}" for i in range(20)],
        )
        large_height = panel.height()
        self.assertGreater(large_height, small_height)
        self.assertLessEqual(large_height, panel.MAX_HEIGHT)

    def test_variable_rows_are_compact_and_flow_responsively(self):
        panel = VariableInsertPanel()
        panel.set_catalog(
            [f"{{{{系统变量{i}}}}}" for i in range(7)],
            [f"数据列{i}" for i in range(7)],
        )

        for button in panel._sys_buttons + panel._tbl_buttons:
            self.assertEqual(button.sizePolicy().horizontalPolicy(), QSizePolicy.Minimum)

        # The flow contains only natural-width pills
        self.assertEqual(panel.sys_capsules_layout.count(), 7)
        self.assertTrue(all(panel.sys_capsules_layout.itemAt(i).widget() is not None for i in range(7)))
        self.assertTrue(panel.sys_capsules_layout.hasHeightForWidth())

    def test_system_variable_description_is_not_shown(self):
        button = VariableCapsuleButton("{{模板所在文件夹}}")
        self.assertEqual(button.text(), "{{模板所在文件夹}}")
        self.assertNotIn("源目录", button.text())

    def test_create_variable_icon(self):
        icon = create_variable_icon()
        self.assertIsNotNone(icon)
        self.assertFalse(icon.isNull())
