"""Template batch generation page supporting capsule file input, drag-and-drop, embedded variable panel, and 4-step workflow."""

from __future__ import annotations

import dataclasses
import os
from pathlib import Path

from application.merge_service import MergeService
from application.task_coordinator import TaskCoordinator
from application.task_models import ServiceResult, TaskProgress, TaskState
from application.workers import TaskWorker
from core.models import ExcelData, MergePlanResult, MergeResult
from core.template_merge import create_merge_plan, get_table_sheet_names
from platform_adapter.capabilities import CAPABILITIES
from ui.delegates.empty_field_delegate import EmptyFieldDelegate
from ui.delegates.mapping_combo_delegate import MappingComboDelegate
from ui.models.field_mapping_model import FieldMappingModel
from ui.models.template_list_model import TemplateListModel
from ui.theme.theme_manager import THEME, set_theme_tone
from ui.widgets.capsule_edit import DropLineEdit, FileCapsuleEdit
from ui.widgets.card import FluentCard
from ui.widgets.resizable_container import ResizableTableContainer
from ui.widgets.scroll_helper import prevent_wheel_propagation
from ui.widgets.variable_popup import VariableInsertPanel, create_variable_icon

try:
    from PySide6.QtCore import Qt, QThreadPool, Signal
    from PySide6.QtGui import QFontDatabase
    from PySide6.QtWidgets import (
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QTableView,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class MergePage(QWidget):
        """View for multi-template + Excel/CSV batch generation with capsule input, drag-and-drop, and embedded variable panel."""

        statusMessage = Signal(str, str)
        statusMetrics = Signal(str)
        statusProgress = Signal(object)
        statusHideProgress = Signal()

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self._thread_pool = QThreadPool.globalInstance()
            self._active_worker: TaskWorker | None = None
            self._scan_generation_id: int = 0
            self._is_applying_scan_results: bool = False
            self._merge_plan: MergePlanResult | None = None
            self._merge_plan_signature: tuple | None = None

            # State data
            self.excel_data: ExcelData | None = None
            self.template_model = TemplateListModel(parent=self)
            self.mapping_model = FieldMappingModel(parent=self)
            self.combo_delegate = MappingComboDelegate(parent=self)
            self.empty_field_delegate = EmptyFieldDelegate(parent=self)

            # Active target tracking for variable insertion panel
            self._active_target_template_id: str | None = None
            self._active_target_line_edit: QLineEdit | None = None
            self._saved_cursor_pos: int = -1

            self._build_ui()
            self._connect_signals()
            THEME.add_listener(self._apply_theme)
            self._update_step_states()

        def _build_ui(self):
            root_layout = QVBoxLayout(self)
            root_layout.setContentsMargins(0, 0, 0, 0)
            root_layout.setSpacing(0)

            # ---------------- Scroll Area for Content Cards ----------------
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFrameShape(QScrollArea.NoFrame)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.scroll_area.setAccessibleName("模板批量生成内容滚动区")

            self.content_widget = QWidget()
            self.content_widget.setProperty("isPageContent", True)
            layout = QVBoxLayout(self.content_widget)
            layout.setContentsMargins(18, 10, 18, 12)
            layout.setSpacing(12)

            # =========================================================================
            # Step 1: Select Templates & Data Table (Capsule input + Drag & Drop)
            # =========================================================================
            self.card_step1 = FluentCard(
                title="步骤 1：选择模板与数据",
                subtitle="添加待处理的 Word 模板文件（支持拖拽）与 Excel/CSV 表格，然后点击“读取模板与数据”",
            )
            self.card_step1.setAccessibleName("步骤一：选择模板与数据卡片")
            self.card_step1.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            s1_grid = QGridLayout()
            s1_grid.setSpacing(8)

            # Row 0: 模板文件
            lbl_tmpl = QLabel("模板文件：")
            lbl_tmpl.setStyleSheet("font-weight: 500;")
            s1_grid.addWidget(lbl_tmpl, 0, 0)

            self.capsule_templates = FileCapsuleEdit(
                placeholder_text="点击“添加模板…”或将一个/多个 Word 模板文件（.docx/.docm）拖放到此处...",
                supported_extensions=(".docx", ".docm") if CAPABILITIES.is_macos else (".docx", ".docm", ".doc"),
                parent=self,
            )
            self.capsule_templates.setAccessibleName("Word 模板胶囊输入框")
            s1_grid.addWidget(self.capsule_templates, 0, 1)

            self.btn_add_tmpl = QPushButton("添加模板…")
            self.btn_add_tmpl.setCursor(Qt.PointingHandCursor)
            self.btn_add_tmpl.clicked.connect(self._add_templates_dialog)
            s1_grid.addWidget(self.btn_add_tmpl, 0, 2)

            # Row 1: 数据表格
            lbl_excel = QLabel("数据表格：")
            lbl_excel.setStyleSheet("font-weight: 500;")
            s1_grid.addWidget(lbl_excel, 1, 0)

            self.edt_excel = DropLineEdit(
                supported_extensions=(".xlsx", ".csv", ".xlsm", ".xltx", ".xltm"),
                parent=self,
            )
            self.edt_excel.setPlaceholderText("选择或将 .xlsx / .csv 数据表格拖放到此处...")
            self.edt_excel.setAccessibleName("数据源表格路径输入框")
            lbl_excel.setBuddy(self.edt_excel)
            s1_grid.addWidget(self.edt_excel, 1, 1)

            self.btn_browse_excel = QPushButton("添加表格…")
            self.btn_browse_excel.clicked.connect(self._browse_excel)
            s1_grid.addWidget(self.btn_browse_excel, 1, 2)

            # Row 2: 工作表 & 读取按钮
            self.lbl_sheet = QLabel("工作表：")
            self.lbl_sheet.setEnabled(False)
            s1_grid.addWidget(self.lbl_sheet, 2, 0)

            self.combo_sheet = QComboBox()
            self.combo_sheet.setEnabled(False)
            self.combo_sheet.addItem("默认表格")
            self.combo_sheet.setAccessibleName("工作表选择下拉框")
            self.lbl_sheet.setBuddy(self.combo_sheet)
            self.combo_sheet.currentIndexChanged.connect(self._on_sheet_changed)
            s1_grid.addWidget(self.combo_sheet, 2, 1)

            # Large Read Button
            self.btn_scan = QPushButton("读取模板与数据")
            self.btn_scan.setProperty("isPrimary", True)
            self.btn_scan.setAccessibleName("读取模板与数据按钮")
            self.btn_scan.clicked.connect(self.scan_templates)
            s1_grid.addWidget(self.btn_scan, 2, 2)

            # Row 3: Summary label
            self.lbl_source_summary = QLabel("请添加 Word 模板与数据表格后点击“读取模板与数据”。")
            self.lbl_source_summary.setStyleSheet("font-size: 12px;")
            set_theme_tone(self.lbl_source_summary, "secondary")
            s1_grid.addWidget(self.lbl_source_summary, 3, 0, 1, 3)

            self.card_step1.addLayout(s1_grid)
            layout.addWidget(self.card_step1, 0)

            # =========================================================================
            # Step 2: Confirm Field Mapping & Empty Value Policy
            # =========================================================================
            self.card_step2 = FluentCard(
                title="步骤 2：确认字段映射与空值策略",
                subtitle="展示所有模板变量的并集；可下拉修改对应数据列或配置空值替换策略",
            )
            self.card_step2.setAccessibleName("步骤二：字段映射与空值替换卡片")
            self.card_step2.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

            self.tbl_mapping = QTableView()
            self.tbl_mapping.setModel(self.mapping_model)
            self.tbl_mapping.setItemDelegateForColumn(FieldMappingModel.COL_EXCEL, self.combo_delegate)
            self.tbl_mapping.setItemDelegateForColumn(FieldMappingModel.COL_DEFAULT, self.empty_field_delegate)
            self.tbl_mapping.setAccessibleName("字段映射关系表格")

            header = self.tbl_mapping.horizontalHeader()
            header.setSectionResizeMode(FieldMappingModel.COL_VARIABLE, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_EXCEL, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_DEFAULT, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_STATUS, QHeaderView.Fixed)
            header.resizeSection(FieldMappingModel.COL_STATUS, 95)
            self.tbl_mapping.clicked.connect(self._on_mapping_table_clicked)

            self.container_mapping = ResizableTableContainer(
                self.tbl_mapping,
                initial_height=130,
                min_height=80,
                max_height=550,
                parent=self,
            )
            self.card_step2.addWidget(self.container_mapping, 1)

            self.lbl_mapping_summary = QLabel("等待步骤 1 读取模板与数据...")
            self.lbl_mapping_summary.setStyleSheet("font-size: 12px;")
            set_theme_tone(self.lbl_mapping_summary, "secondary")
            self.card_step2.addWidget(self.lbl_mapping_summary, 0)

            layout.addWidget(self.card_step2, 1)

            # =========================================================================
            # Step 3: Configure Output Path & Per-Template Filenames
            # =========================================================================
            self.card_step3 = FluentCard(
                title="步骤 3：设置输出路径与逐模板命名规则",
                subtitle="在下方表格中为每个模板独立配置输出文件夹与文件名规则，点击输入框右侧【{ }】按钮可展开变量面板快捷插入",
            )
            self.card_step3.setAccessibleName("步骤三：输出路径与文件名规则卡片")
            self.card_step3.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            s3_layout = QVBoxLayout()
            s3_layout.setSpacing(8)

            # Template Output Rules Table (Per-template editable full path and filename rules)
            self.tbl_output_templates = QTableView()
            self.tbl_output_templates.setModel(self.template_model)
            self.tbl_output_templates.setAccessibleName("Word 模板输出规则表格")

            out_hdr = self.tbl_output_templates.horizontalHeader()
            out_hdr.setSectionResizeMode(TemplateListModel.COL_ENABLED, QHeaderView.Fixed)
            out_hdr.resizeSection(TemplateListModel.COL_ENABLED, 48)
            out_hdr.setSectionResizeMode(TemplateListModel.COL_NAME, QHeaderView.Interactive)
            out_hdr.resizeSection(TemplateListModel.COL_NAME, 180)
            out_hdr.setSectionResizeMode(TemplateListModel.COL_FILENAME_RULE, QHeaderView.Stretch)

            self.container_output_templates = ResizableTableContainer(
                self.tbl_output_templates,
                initial_height=120,
                min_height=70,
                max_height=500,
                parent=self,
            )
            s3_layout.addWidget(self.container_output_templates)

            # Shared in-page collapsible variable insertion panel (hidden by default)
            self.variable_panel = VariableInsertPanel(parent=self)
            self.variable_panel.variableSelected.connect(self._insert_variable_into_target)
            self.variable_panel.hide()
            s3_layout.addWidget(self.variable_panel)

            # Live directory & file preview label
            self.lbl_live_example = QLabel("目录与文件生成示例：—")
            self.lbl_live_example.setStyleSheet("font-size: 12px; padding: 2px 4px;")
            set_theme_tone(self.lbl_live_example, "secondary")
            s3_layout.addWidget(self.lbl_live_example)

            self.card_step3.addLayout(s3_layout)
            layout.addWidget(self.card_step3, 0)

            # =========================================================================
            # Step 4: Options & Execution Log
            # =========================================================================
            self.card_step4 = FluentCard(
                title="步骤 4：选项与生成日志",
                subtitle="设置选项并在下方查看多模板批量生成预览与实时日志",
            )
            self.card_step4.setAccessibleName("步骤四：选项与生成日志卡片")
            self.card_step4.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            opts_row = QHBoxLayout()
            opts_row.setSpacing(16)

            com_state_hint = "" if CAPABILITIES.has_word_com else "（需 Windows 与 Microsoft Word）"
            self.cb_use_com = QCheckBox(f"启用 Word COM 完整模式{com_state_hint}")
            self.cb_use_com.setEnabled(CAPABILITIES.has_word_com)
            self.cb_use_com.setAccessibleName("启用 Word COM 完整模式复选框")
            opts_row.addWidget(self.cb_use_com)

            opts_row.addStretch()
            self.card_step4.addLayout(opts_row)

            # Log / Output Viewer with Platform Monospace Font
            self.txt_log = QPlainTextEdit()
            self.txt_log.setFixedHeight(95)
            self.txt_log.setReadOnly(True)
            self.txt_log.setPlaceholderText("预览和多模板批量生成的日志将在此处显示...")
            self.txt_log.setAccessibleName("批量生成执行日志区域")

            fixed_font = QFontDatabase.systemFont(QFontDatabase.FixedFont)
            fixed_font.setPointSize(12)
            self.txt_log.setFont(fixed_font)
            prevent_wheel_propagation(self.txt_log)
            self.card_step4.addWidget(self.txt_log)

            layout.addWidget(self.card_step4, 0)

            self.scroll_area.setWidget(self.content_widget)
            root_layout.addWidget(self.scroll_area, 1)

            # ---------------- Fixed Bottom Action Bar (Outside ScrollArea) ----------------
            action_container = QFrame(self)
            action_container.setProperty("isActionBar", True)

            action_bar = QHBoxLayout(action_container)
            action_bar.setContentsMargins(0, 0, 0, 0)
            action_bar.setSpacing(10)

            self.btn_preview = QPushButton("预览生成任务")
            self.btn_preview.setAccessibleName("预览生成任务按钮")
            self.btn_preview.clicked.connect(self.preview_merge)
            action_bar.addWidget(self.btn_preview)

            action_bar.addStretch()

            self.btn_generate = QPushButton("开始批量生成")
            self.btn_generate.setProperty("isPrimary", True)
            self.btn_generate.setAccessibleName("开始批量生成主按钮")
            self.btn_generate.clicked.connect(self.start_batch_merge)
            action_bar.addWidget(self.btn_generate)

            root_layout.addWidget(action_container)

        def _connect_signals(self):
            # Model & Capsule events
            self.capsule_templates.filesDropped.connect(self._on_templates_dropped)
            self.capsule_templates.fileRemoved.connect(self._on_template_capsule_removed)
            self.edt_excel.textChanged.connect(self._on_excel_path_changed)
            self.mapping_model.dataChanged.connect(self._on_mapping_changed)

            self.template_model.dataChanged.connect(self._on_template_data_changed)
            self.template_model.rowsInserted.connect(self._on_templates_structure_changed)
            self.template_model.rowsRemoved.connect(self._on_templates_structure_changed)
            self.template_model.modelReset.connect(self._on_templates_structure_changed)

        def _open_mapping_editor(self, index):
            if index.isValid() and index.column() in (FieldMappingModel.COL_EXCEL, FieldMappingModel.COL_DEFAULT):
                self.tbl_mapping.edit(index)

        # ---------------- State & Event Handlers ----------------

        def _apply_theme(self, _tokens):
            """Refresh dynamic icons and semantic preview colors after a theme change."""
            for row in range(self.template_model.rowCount()):
                idx = self.template_model.index(row, TemplateListModel.COL_FILENAME_RULE)
                line_edit = self.tbl_output_templates.indexWidget(idx)
                if isinstance(line_edit, QLineEdit):
                    action = getattr(line_edit, "_variable_action", None)
                    if action is not None:
                        action.setIcon(create_variable_icon())
            self._update_live_examples()

        def _invalidate_merge_plan(self):
            self._merge_plan = None
            self._merge_plan_signature = None

        def _current_plan_signature(self) -> tuple:
            templates = tuple(
                (t.template_id, t.file_path, t.display_name, t.enabled, t.filename_rule)
                for t in self.template_model.get_templates()
            )
            mapping = tuple(sorted(self.mapping_model.get_mapping().items()))
            defaults = tuple(sorted(self.mapping_model.get_defaults().items()))
            behaviors = tuple(sorted(self.mapping_model.get_empty_behaviors().items()))
            return (id(self.excel_data), templates, mapping, defaults, behaviors)

        def _sync_capsules_from_model(self):
            tmpls = self.template_model.get_templates()
            items = [(t.file_path, t.display_name) for t in tmpls]
            self.capsule_templates.set_files(items)

        def _refresh_embedded_path_inputs(self):
            """Embed native QLineEdit widgets with trailing { } variable button into tbl_output_templates."""
            tmpls = self.template_model.get_templates()
            var_icon = create_variable_icon()

            for row, tmpl in enumerate(tmpls):
                idx = self.template_model.index(row, TemplateListModel.COL_FILENAME_RULE)
                existing = self.tbl_output_templates.indexWidget(idx)
                if isinstance(existing, QLineEdit):
                    if existing.text() != tmpl.filename_rule:
                        existing.blockSignals(True)
                        existing.setText(tmpl.filename_rule)
                        existing.blockSignals(False)
                    continue

                line_edit = QLineEdit(tmpl.filename_rule)
                line_edit.setPlaceholderText("输出路径与文件名规则，例如：D:/输出/{{模板名}}-{{数据序号}}")
                line_edit.setAccessibleName(f"模板【{tmpl.display_name}】输出路径输入框")

                # Trailing action for variable button
                action_var = line_edit.addAction(var_icon, QLineEdit.TrailingPosition)
                action_var.setToolTip(f"为【{tmpl.display_name}】插入变量（展开/折叠变量面板）")
                line_edit._variable_action = action_var

                action_var.triggered.connect(
                    lambda _, tid=tmpl.template_id, le=line_edit: self._toggle_variable_panel(tid, le)
                )

                def _make_handler(target_id):
                    def _on_text_changed(text):
                        for t in self.template_model.get_templates():
                            if t.template_id == target_id:
                                t.filename_rule = text.strip()
                                break
                        self._invalidate_merge_plan()
                        self._update_live_examples()
                    return _on_text_changed

                line_edit.textChanged.connect(_make_handler(tmpl.template_id))
                self.tbl_output_templates.setIndexWidget(idx, line_edit)

        def _toggle_variable_panel(self, template_id: str, line_edit: QLineEdit):
            """Toggle variable panel visibility for the specified template row."""
            if self.variable_panel.isVisible() and self._active_target_template_id == template_id:
                self._close_variable_panel()
            else:
                self._open_variable_panel(template_id, line_edit)

        def _open_variable_panel(self, template_id: str, line_edit: QLineEdit):
            """Open and configure the variable insertion panel for the target template."""
            tmpl = next((t for t in self.template_model.get_templates() if t.template_id == template_id), None)
            if not tmpl:
                return

            self._active_target_template_id = template_id
            self._active_target_line_edit = line_edit
            self._saved_cursor_pos = line_edit.cursorPosition()

            self.variable_panel.set_target_name(tmpl.display_name)
            self.variable_panel.reset_search()
            self.variable_panel.show()

        def _close_variable_panel(self):
            """Hide the variable insertion panel and clear target state."""
            self.variable_panel.hide()
            self._active_target_template_id = None
            self._active_target_line_edit = None
            self._saved_cursor_pos = -1

        def _insert_variable_into_target(self, tag: str):
            """Insert selected variable placeholder into the active target QLineEdit at cursor position."""
            if not self._active_target_template_id:
                self._close_variable_panel()
                return

            tmpl = next((t for t in self.template_model.get_templates() if t.template_id == self._active_target_template_id), None)
            if not tmpl:
                self._close_variable_panel()
                return

            if not self._active_target_line_edit:
                self._close_variable_panel()
                return

            try:
                if self._saved_cursor_pos >= 0:
                    self._active_target_line_edit.setCursorPosition(self._saved_cursor_pos)
                self._active_target_line_edit.insert(tag)
                self._saved_cursor_pos = self._active_target_line_edit.cursorPosition()
                self._active_target_line_edit.setFocus()
            except Exception:
                self._close_variable_panel()

        def _refresh_variable_catalog(self):
            """Update variable catalog in the panel from system variables and Excel headers."""
            sys_vars = [
                "{{模板所在文件夹}}",
                "{{模板名}}",
                "{{模板文件名}}",
                "{{模板序号}}",
                "{{数据序号}}",
                "{{Excel行号}}",
                "{{日期}}",
            ]
            table_vars = list(self.excel_data.headers) if self.excel_data and self.excel_data.headers else []
            self.variable_panel.set_catalog(sys_vars, table_vars)

        def _on_templates_dropped(self, file_paths: list[str]):
            def_dir = str(Path(file_paths[0]).parent / "Generated") if file_paths else ""
            self.template_model.add_templates(file_paths, default_dir=def_dir)

        def _on_template_capsule_removed(self, file_path: str):
            tmpls = self.template_model.get_templates()
            for idx, t in enumerate(tmpls):
                if os.path.normcase(os.path.abspath(t.file_path)) == os.path.normcase(os.path.abspath(file_path)):
                    if self._active_target_template_id == t.template_id:
                        self._close_variable_panel()
                    self.template_model.remove_indices([idx])
                    break

        def _on_mapping_table_clicked(self, index):
            if index.isValid() and index.column() in (FieldMappingModel.COL_EXCEL, FieldMappingModel.COL_DEFAULT):
                self.tbl_mapping.edit(index)

        def _on_template_data_changed(self, top_left, bottom_right, roles=None):
            """Handle in-place edits on existing templates (checkbox toggle or filename rule edit)."""
            self._invalidate_merge_plan()
            self._update_step_states()
            self._update_live_examples()

        def _on_templates_structure_changed(self):
            """Handle structural additions or removals of template files."""
            self._sync_capsules_from_model()
            self._invalidate_merge_plan()
            self._refresh_embedded_path_inputs()
            if not self._is_applying_scan_results:
                self._invalidate_scan_state()
            self._update_step_states()
            self._update_live_examples()

        def _on_excel_path_changed(self):
            self._update_sheet_selector()
            self._invalidate_scan_state()
            self._update_step_states()
            self._update_live_examples()

        def _on_mapping_changed(self):
            self._invalidate_merge_plan()
            self._update_live_examples()

        def _update_sheet_selector(self):
            path = self.edt_excel.text().strip()
            if not path or not os.path.isfile(path):
                self.combo_sheet.blockSignals(True)
                self.combo_sheet.clear()
                self.combo_sheet.addItem("默认表格")
                self.combo_sheet.setEnabled(False)
                self.lbl_sheet.setEnabled(False)
                self.combo_sheet.blockSignals(False)
                return

            sheet_names = get_table_sheet_names(path)
            self.combo_sheet.blockSignals(True)
            self.combo_sheet.clear()
            if len(sheet_names) > 1:
                self.combo_sheet.addItems(sheet_names)
                self.combo_sheet.setEnabled(True)
                self.lbl_sheet.setEnabled(True)
            elif len(sheet_names) == 1:
                self.combo_sheet.addItem(sheet_names[0])
                self.combo_sheet.setEnabled(False)
                self.lbl_sheet.setEnabled(False)
            else:
                self.combo_sheet.addItem("默认表格")
                self.combo_sheet.setEnabled(False)
                self.lbl_sheet.setEnabled(False)
            self.combo_sheet.blockSignals(False)

        def _on_sheet_changed(self, index: int = 0):
            self._invalidate_scan_state()
            self._update_step_states()

        def _invalidate_scan_state(self):
            self._invalidate_merge_plan()
            self._close_variable_panel()
            if self.excel_data is not None:
                self.excel_data = None
                self.mapping_model.set_data([], {})
                self.lbl_source_summary.setText("模板或数据表已变更，请重新点击“读取模板与数据”。")
                set_theme_tone(self.lbl_source_summary, "secondary")
                self.lbl_mapping_summary.setText("等待读取模板与数据...")
                set_theme_tone(self.lbl_mapping_summary, "secondary")

        def _update_step_states(self):
            all_tmpls = self.template_model.get_templates()
            enabled_tmpls = self.template_model.get_enabled_templates()
            e_path = self.edt_excel.text().strip()
            can_scan = bool(all_tmpls and e_path and os.path.isfile(e_path))
            self.btn_scan.setEnabled(can_scan)

            has_scanned = self.excel_data is not None and len(self.mapping_model.get_fields()) > 0
            self.card_step2.setEnabled(has_scanned)
            self.card_step3.setEnabled(has_scanned)
            self.card_step4.setEnabled(has_scanned)
            self.btn_preview.setEnabled(has_scanned and bool(enabled_tmpls))
            self.btn_generate.setEnabled(has_scanned and bool(enabled_tmpls))

        def _update_live_examples(self):
            """Render real-time sample destination paths in Step 3 based on row 1 data."""
            if not self.excel_data or not self.excel_data.rows:
                self.lbl_live_example.setText("目录与文件生成示例：—")
                return

            enabled_tmpls = self.template_model.get_enabled_templates()
            if not enabled_tmpls:
                self.lbl_live_example.setText("目录与文件生成示例：请至少在上方勾选启用一个 Word 模板")
                return

            mapping = self.mapping_model.get_mapping()
            defaults = self.mapping_model.get_defaults()
            empty_behaviors = self.mapping_model.get_empty_behaviors()

            try:
                sample_data = ExcelData(
                    headers=list(self.excel_data.headers),
                    rows=[dict(self.excel_data.rows[0])],
                    excel_rows=[self.excel_data.excel_rows[0] if self.excel_data.excel_rows else 2],
                )
                sample_plan = create_merge_plan(
                    templates=[dataclasses.replace(t) for t in enabled_tmpls[:3]],
                    data=sample_data,
                    mapping=mapping,
                    default_values=defaults,
                    empty_field_behaviors=empty_behaviors,
                    check_existing=False,
                )
                sample_lines = [
                    f"   【{job.template_name}】→ {job.destination}"
                    for job in sample_plan.jobs
                ]

                if len(enabled_tmpls) > 3:
                    sample_lines.append(f"   ...以及其他 {len(enabled_tmpls) - 3} 个模板")

                self.lbl_live_example.setText("💡 实时生成示例（基于第 1 行数据）：\n" + "\n".join(sample_lines))
                self.lbl_live_example.setStyleSheet("font-size: 12px; font-weight: 500;")
                set_theme_tone(self.lbl_live_example, "accent")
            except Exception as exc:
                self.lbl_live_example.setText(f"⚠️ 规则提示：{exc}")
                self.lbl_live_example.setStyleSheet("font-size: 12px;")
                set_theme_tone(self.lbl_live_example, "error")

        def _add_templates_dialog(self):
            filter_str = (
                "Word 模板 (*.docx *.docm);;所有文件 (*.*)"
                if CAPABILITIES.is_macos
                else "Word 模板 (*.docx *.docm *.doc);;所有文件 (*.*)"
            )
            paths, _ = QFileDialog.getOpenFileNames(self, "选择一个或多个 Word 模板", "", filter_str)
            if paths:
                def_dir = str(Path(paths[0]).parent / "Generated")
                self.template_model.add_templates(paths, default_dir=def_dir)
                self._update_step_states()
                self._update_live_examples()

        def _browse_excel(self):
            filter_str = (
                "表格数据文件 (*.xlsx *.csv *.xlsm);;"
                "Excel 表格 (*.xlsx *.xlsm);;"
                "CSV 文本表格 (*.csv);;"
                "所有文件 (*.*)"
            )
            path, _ = QFileDialog.getOpenFileName(
                self, "选择数据源表格 (Excel / CSV)", "", filter_str
            )
            if path:
                self.edt_excel.setText(os.path.abspath(path))

        # ---------------- Actions ----------------

        def scan_templates(self):
            all_tmpls = self.template_model.get_templates()
            e_path = self.edt_excel.text().strip()
            use_com = self.cb_use_com.isChecked()

            if not all_tmpls:
                QMessageBox.warning(self, "提示", "请先添加至少一个 Word 模板。")
                return

            paths = [t.file_path for t in all_tmpls]
            err = MergeService.validate_sources(paths, e_path, use_com=use_com)
            if err:
                QMessageBox.warning(self, "提示", err)
                return

            coordinator = TaskCoordinator.instance()
            if not coordinator.can_start_task(is_write=False):
                QMessageBox.warning(self, "提示", "已有后台任务正在执行中，请稍后再试。")
                return

            self._scan_generation_id += 1
            current_scan_gen = self._scan_generation_id

            selected_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None

            self.statusMessage.emit("正在读取各模板与数据表...", "running")
            self._set_scan_inputs_locked(True)

            tmpl_snapshots = [dataclasses.replace(t) for t in all_tmpls]

            def _scan_worker_fn(tmpls, excel_file, sheet, com_flag, progress_cb=None, cancel_token=None):
                res = MergeService.scan_templates_and_excel(
                    templates=tmpls,
                    excel_path=excel_file,
                    use_com=com_flag,
                    sheet_name=sheet,
                    progress_cb=progress_cb,
                    cancel_token=cancel_token,
                )
                return (res, current_scan_gen, excel_file, sheet)

            worker = TaskWorker(
                _scan_worker_fn,
                tmpls=tmpl_snapshots,
                excel_file=e_path,
                sheet=selected_sheet,
                com_flag=use_com,
            )
            self._active_worker = worker
            task_id = coordinator.register_task(worker, is_write=False, description="读取模板与数据")

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_scan_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: (self._set_scan_inputs_locked(False), coordinator.release_task(task_id)))

            self._thread_pool.start(worker)

        def _on_scan_finished(self, result_tuple: tuple[ServiceResult, int, str, str | None]):
            self.statusHideProgress.emit()
            res, gen_id, scanned_excel, scanned_sheet = result_tuple

            if gen_id != self._scan_generation_id:
                return

            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("读取已取消", "warning")
                    return
                self.statusMessage.emit("读取失败", "error")
                QMessageBox.critical(self, "读取失败", res.error)
                return

            scanned_tmpls, all_fields, field_tmpls, excel_data, mapping = res.data
            self.excel_data = excel_data
            self._is_applying_scan_results = True
            try:
                self.template_model.set_scan_results(scanned_tmpls)
                self.combo_delegate.set_headers(excel_data.headers)
                self.mapping_model.set_data(all_fields, mapping, field_templates=field_tmpls)
            finally:
                self._is_applying_scan_results = False

            missing = self.mapping_model.get_missing_count()
            excel_cols_str = "、".join(excel_data.headers[:8])
            if len(excel_data.headers) > 8:
                excel_cols_str += " 等"
            sheet_info = f"（工作表：{scanned_sheet}）" if scanned_sheet else ""
            summary_s1 = f"已读取 {len(scanned_tmpls)} 个模板，{len(all_fields)} 个变量并集，{len(excel_data.rows)} 行数据。"
            self.lbl_source_summary.setText(summary_s1)
            is_warning = (res.state == TaskState.WARNING)
            self.lbl_source_summary.setStyleSheet("font-weight: 500;")
            set_theme_tone(self.lbl_source_summary, "warning" if is_warning else "success")

            summary_s2 = (
                f"变量并集确认{sheet_info}：共检测到 {len(all_fields)} 个模板变量；已匹配 {len(all_fields) - missing} 个，未匹配 {missing} 个。\n"
                f"💡 可用数据列：{excel_cols_str}"
            )
            self.lbl_mapping_summary.setText(summary_s2)
            self.lbl_mapping_summary.setStyleSheet("font-weight: 500;")
            set_theme_tone(
                self.lbl_mapping_summary,
                "error" if missing > 0 else ("warning" if is_warning else "success"),
            )

            self.statusMessage.emit("读取完成（部分模板存在警告或失败）" if is_warning else "读取完成", "warning" if is_warning else "success")
            self.statusMetrics.emit(f"已加载 {len(scanned_tmpls)} 个模板，{len(excel_data.rows)} 行数据")
            self._refresh_variable_catalog()
            self._update_step_states()
            self._update_live_examples()

        def _set_scan_inputs_locked(self, locked: bool):
            self.btn_scan.setEnabled(not locked)
            self.btn_add_tmpl.setEnabled(not locked)
            self.capsule_templates.setEnabled(not locked)
            self.edt_excel.setEnabled(not locked)
            self.btn_browse_excel.setEnabled(not locked)

        def preview_merge(self):
            if not self.excel_data:
                return

            coordinator = TaskCoordinator.instance()
            if not coordinator.can_start_task(is_write=False):
                QMessageBox.warning(self, "提示", "已有任务正在后台执行中，请稍后。")
                return

            enabled_tmpls = self.template_model.get_enabled_templates()
            mapping = self.mapping_model.get_mapping()
            defaults = self.mapping_model.get_defaults()
            empty_behaviors = self.mapping_model.get_empty_behaviors()

            for t in enabled_tmpls:
                if not t.filename_rule or not t.filename_rule.strip():
                    QMessageBox.warning(self, "提示", f"模板【{t.display_name}】的输出路径与文件名规则不能为空。")
                    return

            tmpl_snapshots = [dataclasses.replace(t) for t in enabled_tmpls]
            signature = self._current_plan_signature()
            self._set_ui_busy(True)
            self.statusMessage.emit("正在规划生成任务...", "running")

            worker = TaskWorker(
                MergeService.prepare_merge_preview,
                templates=tmpl_snapshots,
                excel_data=self.excel_data,
                mapping=mapping,
                output_directory_rule="",
                default_values=defaults,
                empty_field_behaviors=empty_behaviors,
                max_rows=5,
            )
            self._active_worker = worker
            task_id = coordinator.register_task(worker, is_write=False, description="规划模板生成任务")
            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(lambda res, sig=signature: self._on_preview_finished(res, sig))
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: (self._set_ui_busy(False), coordinator.release_task(task_id)))
            self._thread_pool.start(worker)

        def _on_preview_finished(self, res: ServiceResult, signature: tuple):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("任务预览已取消", "warning")
                    return
                self.statusMessage.emit("任务规划失败", "error")
                QMessageBox.critical(self, "预览失败", res.error)
                return
            if signature != self._current_plan_signature():
                self.statusMessage.emit("输入已变更，本次预览已作废", "warning")
                return

            plan, snippets = res.data
            self._merge_plan = plan
            self._merge_plan_signature = signature
            enabled_count = len(self.template_model.get_enabled_templates())
            total_jobs = len(plan.jobs)

            lines = [
                "多模板批量生成任务规划预览",
                "=" * 60,
                f"启用模板数：{enabled_count} 个",
                f"数据行数：{len(self.excel_data.rows)} 行",
                f"预计生成文件总数：{total_jobs} 份",
                "=" * 60,
                "",
            ]

            for snip in snippets:
                rel_desc = f"{snip['relative_folder']}/" if snip['relative_folder'] else ""
                collision_hint = f"（{snip['collision_reason']}，已自动编号）" if snip.get('collision_reason') else ""
                lines.append(f"【{snip['template_name']}】第 {snip['excel_row']} 行 => 输出目标：{rel_desc}{snip['filename']} {collision_hint}")
                for f_info in snip["fields"]:
                    lines.append(f"   {{{{{f_info['field']}}}}} => {f_info['value']}")
                lines.append("")

            self.txt_log.setPlainText("\n".join(lines))
            self.statusMessage.emit("任务预览就绪", "warning" if plan.issues else "normal")

        def start_batch_merge(self):
            if not self.excel_data:
                return

            coordinator = TaskCoordinator.instance()
            if not coordinator.can_start_task(is_write=True):
                QMessageBox.warning(self, "提示", "已有任务正在后台执行中，请稍后。")
                return

            enabled_tmpls = self.template_model.get_enabled_templates()

            if not enabled_tmpls:
                QMessageBox.warning(self, "提示", "请至少启用一个 Word 模板。")
                return

            for t in enabled_tmpls:
                if not t.filename_rule or not t.filename_rule.strip():
                    QMessageBox.warning(self, "提示", f"模板【{t.display_name}】的输出路径与文件名规则不能为空。")
                    return

            mapping = self.mapping_model.get_mapping()
            defaults = self.mapping_model.get_defaults()
            empty_behaviors = self.mapping_model.get_empty_behaviors()
            missing = [
                f"{{{{{f}}}}}"
                for f, h in mapping.items()
                if not h and empty_behaviors.get(f) == FieldMappingModel.EMPTY_KEEP
            ]
            if missing:
                reply = QMessageBox.question(
                    self,
                    "存在未匹配变量",
                    "以下模板变量未匹配数据列，并设置为保持原变量：\n\n"
                    + "\n".join(missing)
                    + "\n\n是否继续批量生成？",
                )
                if reply != QMessageBox.Yes:
                    return

            self._set_ui_busy(True)
            self.statusMessage.emit("正在批量生成多模板文档...", "running")
            self.txt_log.setPlainText("正在开始多模板批量生成任务...\n")

            tmpl_snapshots = [dataclasses.replace(t) for t in enabled_tmpls]
            reusable_plan = (
                self._merge_plan
                if self._merge_plan_signature == self._current_plan_signature()
                else None
            )

            worker = TaskWorker(
                MergeService.execute_multi_template_merge,
                templates=tmpl_snapshots,
                excel_data=self.excel_data,
                mapping=mapping,
                output_directory_rule="",
                use_com=self.cb_use_com.isChecked(),
                default_values=defaults,
                empty_field_behaviors=empty_behaviors,
                plan=reusable_plan,
            )
            self._active_worker = worker
            task_id = coordinator.register_task(worker, is_write=True, description="批量生成多模板文档")

            worker.signals.progress.connect(self._on_merge_progress)
            worker.signals.result.connect(self._on_merge_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: (self._set_ui_busy(False), coordinator.release_task(task_id)))

            self._thread_pool.start(worker)

        def _on_merge_progress(self, progress: TaskProgress):
            self.statusProgress.emit(progress)
            if progress.data and isinstance(progress.data, MergeResult):
                r = progress.data
                status_str = "✓ 成功" if r.success else f"✕ 失败 ({r.error})"
                rel_desc = f"{r.relative_folder}/" if r.relative_folder else ""
                self.txt_log.appendPlainText(f"[{r.template_name}] 第 {r.excel_row} 行 | {rel_desc}{r.filename} | {status_str}")

        def _on_merge_finished(self, res: ServiceResult):
            self.statusHideProgress.emit()
            self._invalidate_merge_plan()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("批量生成已取消", "warning")
                    self.txt_log.appendPlainText("\n任务已由用户取消。")
                    return
                self.statusMessage.emit("生成失败", "error")
                QMessageBox.critical(self, "生成失败", res.error)
                return

            results: list[MergeResult] = res.data or []
            success_count = sum(1 for r in results if r.success)
            fail_count = len(results) - success_count

            summary = f"多模板批量生成完成：共生成 {len(results)} 份文档（成功 {success_count} 份，失败 {fail_count} 份）。"
            self.statusMessage.emit("生成完成", "success" if fail_count == 0 else "warning")
            QMessageBox.information(self, "生成完成", summary)

        def _on_task_error(self, err_msg: str):
            self.statusHideProgress.emit()
            self.statusMessage.emit("任务异常", "error")
            QMessageBox.critical(self, "执行异常", err_msg)

        def cancel_active_task(self):
            if self._active_worker:
                self._active_worker.cancel()
                self.statusMessage.emit("正在请求取消任务...", "warning")

        def _set_ui_busy(self, is_busy: bool):
            self.btn_scan.setEnabled(not is_busy)
            self.btn_preview.setEnabled(not is_busy)
            self.btn_generate.setEnabled(not is_busy)
            self.btn_add_tmpl.setEnabled(not is_busy)
            self.capsule_templates.setEnabled(not is_busy)
            if not is_busy:
                self._active_worker = None
                self._update_step_states()

else:

    class MergePage:  # type: ignore
        pass
