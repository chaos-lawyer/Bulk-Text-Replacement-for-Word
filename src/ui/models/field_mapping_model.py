"""Qt model for template mappings and per-field empty-value behavior."""

from __future__ import annotations

try:
    from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
    from PySide6.QtGui import QBrush, QColor

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FieldMappingModel(QAbstractTableModel):
        """Manage Word mappings and the empty-value policy for each field."""

        COL_VARIABLE = 0
        COL_EXCEL = 1
        COL_DEFAULT = 2
        COL_STATUS = 3

        EMPTY_KEEP = "keep_variable"
        EMPTY_REPLACE = "replace_empty"
        EMPTY_CUSTOM = "custom"
        VALID_EMPTY_BEHAVIORS = {EMPTY_KEEP, EMPTY_REPLACE, EMPTY_CUSTOM}

        HEADERS = ["Word 模板变量", "对应数据列（可下拉修改）", "空字段处理", "匹配状态"]

        def __init__(
            self,
            fields: list[str] | None = None,
            mapping: dict[str, str] | None = None,
            defaults: dict[str, str] | None = None,
            field_templates: dict[str, list[str]] | None = None,
            parent=None,
        ):
            super().__init__(parent)
            self._fields: list[str] = list(fields) if fields else []
            self._mapping: dict[str, str] = dict(mapping) if mapping else {}
            self._defaults: dict[str, str] = dict(defaults) if defaults else {}
            self._field_templates: dict[str, list[str]] = dict(field_templates) if field_templates else {}
            self._empty_behaviors: dict[str, str] = {
                field: self.EMPTY_CUSTOM if field in self._defaults else self.EMPTY_KEEP
                for field in self._fields
            }

        def rowCount(self, parent=QModelIndex()) -> int:
            return len(self._fields)

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

            base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
            if index.column() in (self.COL_EXCEL, self.COL_DEFAULT):
                return base_flags | Qt.ItemIsEditable
            return base_flags

        def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
            if not index.isValid() or not (0 <= index.row() < len(self._fields)):
                return None

            field = self._fields[index.row()]
            excel_col = self._mapping.get(field, "")
            is_matched = bool(excel_col)
            default_val = self._defaults.get(field, "")
            empty_behavior = self._empty_behaviors.get(field, self.EMPTY_KEEP)
            col = index.column()

            if role == Qt.DisplayRole:
                if col == self.COL_VARIABLE:
                    return f"{{{{{field}}}}}"
                elif col == self.COL_EXCEL:
                    return excel_col if excel_col else "（未指定）"
                elif col == self.COL_DEFAULT:
                    if empty_behavior == self.EMPTY_REPLACE:
                        return "替换为空"
                    if empty_behavior == self.EMPTY_CUSTOM:
                        return f"自定义：{default_val}" if default_val else "自定义…"
                    return "保持原变量（默认）"
                elif col == self.COL_STATUS:
                    return "✓ 已匹配" if is_matched else "未匹配"

            elif role == Qt.EditRole:
                if col == self.COL_EXCEL:
                    return excel_col
                elif col == self.COL_DEFAULT:
                    return (empty_behavior, default_val)

            elif role == Qt.ToolTipRole:
                if col == self.COL_VARIABLE:
                    tmpls = self._field_templates.get(field, [])
                    tips = [f"模板变量：{{{{{field}}}}}"]
                    if tmpls:
                        tips.append("所属模板：\n" + "\n".join(f"• {t}" for t in tmpls))
                    return "\n".join(tips)
                elif col == self.COL_DEFAULT and empty_behavior == self.EMPTY_CUSTOM and default_val:
                    return f"空字段替补文本：\n{default_val}"

            elif role == Qt.ForegroundRole:
                if col == self.COL_STATUS:
                    return QBrush(QColor("#0F7B0F" if is_matched else "#C42B1C"))

            return None

        def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
            if not index.isValid() or not (0 <= index.row() < len(self._fields)):
                return False

            field = self._fields[index.row()]
            col = index.column()

            if role == Qt.EditRole:
                if col == self.COL_EXCEL:
                    self._mapping[field] = str(value)
                    self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                    status_idx = self.index(index.row(), self.COL_STATUS)
                    self.dataChanged.emit(status_idx, status_idx, [Qt.DisplayRole, Qt.ForegroundRole])
                    return True
                elif col == self.COL_DEFAULT:
                    if isinstance(value, tuple) and len(value) == 2:
                        behavior, custom_val = value
                        if behavior in self.VALID_EMPTY_BEHAVIORS:
                            self._empty_behaviors[field] = behavior
                        if custom_val:
                            self._defaults[field] = str(custom_val)
                        elif field in self._defaults and behavior != self.EMPTY_CUSTOM:
                            del self._defaults[field]
                    elif isinstance(value, str):
                        if value in self.VALID_EMPTY_BEHAVIORS:
                            self._empty_behaviors[field] = value
                            if value != self.EMPTY_CUSTOM and field in self._defaults:
                                del self._defaults[field]
                        else:
                            if value:
                                self._defaults[field] = value
                                self._empty_behaviors[field] = self.EMPTY_CUSTOM
                            elif field in self._defaults:
                                del self._defaults[field]

                    self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                    return True

            return False

        def set_data(
            self,
            fields: list[str],
            mapping: dict[str, str],
            field_templates: dict[str, list[str]] | None = None,
        ) -> None:
            """Preserve previous custom defaults and behaviors when re-scanning."""
            self.beginResetModel()
            prev_defaults = dict(self._defaults)
            prev_behaviors = dict(self._empty_behaviors)

            self._fields = list(fields)
            self._mapping = dict(mapping)
            self._field_templates = dict(field_templates) if field_templates else {}

            self._defaults = {f: prev_defaults[f] for f in self._fields if f in prev_defaults}
            self._empty_behaviors = {
                f: prev_behaviors.get(f, self.EMPTY_CUSTOM if f in self._defaults else self.EMPTY_KEEP)
                for f in self._fields
            }
            self.endResetModel()

        def get_fields(self) -> list[str]:
            return list(self._fields)

        def get_mapping(self) -> dict[str, str]:
            return dict(self._mapping)

        def get_defaults(self) -> dict[str, str]:
            return dict(self._defaults)

        def get_empty_behaviors(self) -> dict[str, str]:
            return dict(self._empty_behaviors)

        def get_missing_count(self) -> int:
            return sum(1 for f in self._fields if not self._mapping.get(f))

else:

    class FieldMappingModel:  # type: ignore
        pass
