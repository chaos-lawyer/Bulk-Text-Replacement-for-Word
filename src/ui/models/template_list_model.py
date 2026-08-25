"""Qt Table Models for managing multiple Word templates with source management and per-template naming rules."""

from __future__ import annotations

import os
from pathlib import Path
import uuid

from core.models import TemplateMergeItem

try:
    from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class TemplateListModel(QAbstractTableModel):
        """Table model representing Word templates with 3 columns: Enabled, Display Name, and Output Path/Filename Rule."""

        COL_ENABLED = 0
        COL_NAME = 1
        COL_FILENAME_RULE = 2

        HEADERS = ["启用", "模板名称", "输出路径与文件名规则"]

        def __init__(self, parent=None):
            super().__init__(parent)
            self._items: list[TemplateMergeItem] = []

        def rowCount(self, parent=QModelIndex()) -> int:
            return len(self._items)

        def columnCount(self, parent=QModelIndex()) -> int:
            return len(self.HEADERS)

        def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
            if orientation == Qt.Horizontal and role == Qt.DisplayRole:
                if 0 <= section < len(self.HEADERS):
                    return self.HEADERS[section]
            return None

        def flags(self, index: QModelIndex) -> Qt.ItemFlags:
            if not index.isValid():
                return Qt.NoItemFlags

            col = index.column()
            base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable

            if col == self.COL_ENABLED:
                return base_flags | Qt.ItemIsUserCheckable
            elif col in (self.COL_NAME, self.COL_FILENAME_RULE):
                return base_flags | Qt.ItemIsEditable
            return base_flags

        def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
            if not index.isValid() or not (0 <= index.row() < len(self._items)):
                return None

            item = self._items[index.row()]
            col = index.column()

            if role == Qt.CheckStateRole and col == self.COL_ENABLED:
                return Qt.Checked if item.enabled else Qt.Unchecked

            if role in (Qt.DisplayRole, Qt.EditRole):
                if col == self.COL_NAME:
                    return item.display_name
                elif col == self.COL_FILENAME_RULE:
                    return item.filename_rule

            if role == Qt.ToolTipRole:
                if col == self.COL_NAME:
                    tips = [f"文件路径：{item.file_path}"]
                    if item.fields:
                        vars_str = "、".join(f"{{{{{f}}}}}" for f in item.fields)
                        tips.append(f"检测变量（{len(item.fields)}个）：{vars_str}")
                    if item.error:
                        tips.append(f"扫描异常：{item.error}")
                    elif item.status and item.status != "待扫描":
                        tips.append(f"状态：{item.status}")
                    return "\n".join(tips)
                elif col == self.COL_FILENAME_RULE:
                    return f"输出完整路径与文件名规则：\n{item.filename_rule}\n（可直接输入修改文件夹与文件名）"

            return None

        def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
            if not index.isValid() or not (0 <= index.row() < len(self._items)):
                return False

            item = self._items[index.row()]
            col = index.column()

            if role == Qt.CheckStateRole and col == self.COL_ENABLED:
                item.enabled = (value == Qt.Checked or value == 2 or value is True)
                self.dataChanged.emit(index, index, [Qt.CheckStateRole])
                return True

            if role == Qt.EditRole:
                if col == self.COL_NAME:
                    new_name = str(value).strip()
                    if new_name and new_name != item.display_name:
                        item.display_name = new_name
                        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                        return True
                elif col == self.COL_FILENAME_RULE:
                    new_rule = str(value).strip()
                    if new_rule != item.filename_rule:
                        item.filename_rule = new_rule
                        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                        return True

            return False

        def add_templates(self, paths: list[str], default_dir: str = "") -> int:
            """Add new template paths with default output path rule, deduplicating by normalized path."""
            existing_paths = {os.path.normcase(os.path.abspath(i.file_path)) for i in self._items}
            new_items: list[TemplateMergeItem] = []

            for p in paths:
                norm_p = os.path.normcase(os.path.abspath(p))
                if norm_p not in existing_paths:
                    tmpl_id = f"tmpl_{uuid.uuid4().hex[:8]}"
                    clean_dir = default_dir.strip().rstrip("/\\") if default_dir else ""
                    default_rule = f"{clean_dir}/{{{{模板名}}}}-{{{{数据序号}}}}" if clean_dir else "{{模板名}}-{{数据序号}}"
                    item = TemplateMergeItem(
                        template_id=tmpl_id,
                        file_path=os.path.abspath(p),
                        display_name=Path(p).stem,
                        filename_rule=default_rule,
                        enabled=True,
                        status="待扫描",
                    )
                    new_items.append(item)
                    existing_paths.add(norm_p)

            if not new_items:
                return 0

            start_row = len(self._items)
            end_row = start_row + len(new_items) - 1
            self.beginInsertRows(QModelIndex(), start_row, end_row)
            self._items.extend(new_items)
            self.endInsertRows()
            return len(new_items)

        def remove_indices(self, rows: list[int]) -> None:
            """Remove templates at specified row indices."""
            for row in sorted(rows, reverse=True):
                if 0 <= row < len(self._items):
                    self.beginRemoveRows(QModelIndex(), row, row)
                    del self._items[row]
                    self.endRemoveRows()

        def clear(self) -> None:
            """Clear all templates."""
            self.beginResetModel()
            self._items.clear()
            self.endResetModel()

        def get_templates(self) -> list[TemplateMergeItem]:
            return list(self._items)

        def get_enabled_templates(self) -> list[TemplateMergeItem]:
            return [t for t in self._items if t.enabled]

        def set_scan_results(self, scanned_templates: list[TemplateMergeItem]) -> None:
            """Update template scan states and fields in-place, preserving user edited names & rules."""
            self.beginResetModel()
            lookup = {t.template_id: t for t in scanned_templates}
            for item in self._items:
                if item.template_id in lookup:
                    scanned = lookup[item.template_id]
                    item.fields = list(scanned.fields)
                    item.status = scanned.status
                    item.error = scanned.error
            self.endResetModel()

else:

    class TemplateListModel:  # type: ignore
        pass
