import unittest
from pathlib import Path
import tempfile

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QLineEdit
    from ui.pages.merge_page import MergePage
    from ui.models.template_list_model import TemplateListModel

    app = QApplication.instance() or QApplication([])
    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 not available")
class MergePageVariablePanelIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.page = MergePage()
        self.page.show()

    def tearDown(self):
        self.page.close()
        self.page.deleteLater()

    def test_variable_panel_toggle_and_insertion(self):
        # Add a template
        self.page.template_model.add_templates(["/dummy/path/采购合同.docx"], default_dir="D:/Output")
        tmpls = self.page.template_model.get_templates()
        self.assertEqual(len(tmpls), 1)
        tmpl = tmpls[0]

        # Ensure embedded QLineEdit is created
        idx = self.page.template_model.index(0, TemplateListModel.COL_FILENAME_RULE)
        line_edit = self.page.tbl_output_templates.indexWidget(idx)
        self.assertIsInstance(line_edit, QLineEdit)
        self.assertEqual(line_edit.text(), tmpl.filename_rule)

        # Panel is initially hidden
        self.assertFalse(self.page.variable_panel.isVisible())

        # Toggle open panel for this template
        self.page._open_variable_panel(tmpl.template_id, line_edit)
        self.assertTrue(self.page.variable_panel.isVisible())
        self.assertIn("采购合同", self.page.variable_panel.lbl_target.text())

        # Place cursor at end and insert variable
        line_edit.setCursorPosition(len(line_edit.text()))
        self.page._insert_variable_into_target("{{数据序号}}")

        # Verify inserted into QLineEdit and updated model
        self.assertIn("{{数据序号}}", line_edit.text())
        self.assertIn("{{数据序号}}", tmpl.filename_rule)

        # Panel remains open for multi-insertion
        self.assertTrue(self.page.variable_panel.isVisible())

        # Close panel
        self.page._close_variable_panel()
        self.assertFalse(self.page.variable_panel.isVisible())

    def test_toggle_enabled_checkbox_does_not_clear_step2(self):
        from core.models import ExcelData

        # Simulate completed scan state
        self.page.template_model.add_templates(["/dummy/path/采购合同.docx", "/dummy/path/保密协议.docx"])
        self.page.excel_data = ExcelData(headers=["客户姓名", "合同金额"], rows=[{"客户姓名": "张三", "合同金额": "100"}], excel_rows=[2])
        self.page.mapping_model.set_data(["客户姓名", "合同金额"], {"客户姓名": "客户姓名", "合同金额": "合同金额"})

        self.assertEqual(len(self.page.mapping_model.get_fields()), 2)
        self.assertIsNotNone(self.page.excel_data)

        # Toggle enabled checkbox on template 0 (uncheck it)
        idx = self.page.template_model.index(0, TemplateListModel.COL_ENABLED)
        self.page.template_model.setData(idx, Qt.Unchecked, Qt.CheckStateRole)

        # Verify Step 2 mapping is NOT cleared!
        self.assertEqual(len(self.page.mapping_model.get_fields()), 2)
        self.assertIsNotNone(self.page.excel_data)
        self.assertEqual(self.page.mapping_model.get_mapping()["客户姓名"], "客户姓名")

        # Edit filename rule on template 1
        rule_idx = self.page.template_model.index(1, TemplateListModel.COL_FILENAME_RULE)
        self.page.template_model.setData(rule_idx, "D:/Output/CustomRule", Qt.EditRole)

        # Verify Step 2 mapping is still NOT cleared!
        self.assertEqual(len(self.page.mapping_model.get_fields()), 2)
        self.assertIsNotNone(self.page.excel_data)

    def test_preview_builds_reusable_plan_and_rule_edit_invalidates_it(self):
        from core.models import ExcelData
        from application.task_coordinator import TaskCoordinator

        with tempfile.TemporaryDirectory() as folder:
            template_path = Path(folder) / "template.docx"
            doc = Document()
            doc.add_paragraph("客户：{{客户姓名}}")
            doc.save(template_path)

            self.page.template_model.add_templates(
                [str(template_path)],
                default_dir=str(Path(folder) / "Output"),
            )
            self.page.excel_data = ExcelData(
                headers=["客户姓名"],
                rows=[{"客户姓名": "张三"}],
                excel_rows=[2],
            )
            self.page.mapping_model.set_data(["客户姓名"], {"客户姓名": "客户姓名"})
            self.page._update_step_states()

            self.page.preview_merge()
            self.page._thread_pool.waitForDone(5000)
            app.processEvents()

            self.assertIsNotNone(self.page._merge_plan)
            self.assertEqual(len(self.page._merge_plan.jobs), 1)
            self.assertFalse(TaskCoordinator.instance().is_busy)

            idx = self.page.template_model.index(0, TemplateListModel.COL_FILENAME_RULE)
            line_edit = self.page.tbl_output_templates.indexWidget(idx)
            line_edit.setText(str(Path(folder) / "Output" / "changed-{{客户姓名}}"))
            self.assertIsNone(self.page._merge_plan)
