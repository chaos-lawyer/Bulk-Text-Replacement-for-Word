"""Embedded variable insertion panel and capsule buttons for template batch generation."""

from __future__ import annotations

from ui.theme.theme_manager import THEME
from ui.theme.tokens import ColorTokens
from ui.widgets.scroll_helper import prevent_wheel_propagation

try:
    from PySide6.QtCore import QPoint, QRect, QSize, Qt, Signal
    from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
    from PySide6.QtWidgets import (
        QFrame,
        QHBoxLayout,
        QLabel,
        QLayout,
        QLayoutItem,
        QLineEdit,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FlowLayout(QLayout):
        """A responsive flow layout packing widgets with natural wrapping."""

        def __init__(self, parent: QWidget | None = None, margin: int = 0, h_spacing: int = 6, v_spacing: int = 6):
            super().__init__(parent)
            self._h_spacing = h_spacing
            self._v_spacing = v_spacing
            self._items: list[QLayoutItem] = []
            self.setContentsMargins(margin, margin, margin, margin)

        def addItem(self, item: QLayoutItem):
            self._items.append(item)

        def count(self):
            return len(self._items)

        def itemAt(self, index):
            return self._items[index] if 0 <= index < len(self._items) else None

        def takeAt(self, index):
            return self._items.pop(index) if 0 <= index < len(self._items) else None

        def expandingDirections(self):
            return Qt.Orientations(Qt.Orientation(0))

        def hasHeightForWidth(self):
            return True

        def heightForWidth(self, width: int):
            return self._do_layout(QRect(0, 0, width, 0), True)

        def setGeometry(self, rect: QRect):
            super().setGeometry(rect)
            self._do_layout(rect, False)

        def sizeHint(self):
            return self.minimumSize()

        def minimumSize(self):
            size = QSize()
            for item in self._items:
                size = size.expandedTo(item.minimumSize())
            margins = self.contentsMargins()
            size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
            return size

        def _do_layout(self, rect: QRect, test_only: bool) -> int:
            margins = self.contentsMargins()
            effective = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
            x = effective.x()
            y = effective.y()
            line_height = 0

            for item in self._items:
                widget = item.widget()
                if widget is not None and widget.isHidden():
                    continue

                hint = item.sizeHint()
                next_x = x + hint.width() + self._h_spacing
                if line_height > 0 and next_x - self._h_spacing > effective.right() + 1:
                    x = effective.x()
                    y += line_height + self._v_spacing
                    next_x = x + hint.width() + self._h_spacing
                    line_height = 0

                if not test_only:
                    item.setGeometry(QRect(QPoint(x, y), hint))

                x = next_x
                line_height = max(line_height, hint.height())

            return y + line_height - rect.y() + margins.bottom()


    class VariableCapsuleButton(QPushButton):
        """Pill-shaped clickable capsule representing a placeholder variable."""

        def __init__(self, tag: str, parent: QWidget | None = None):
            self.tag = tag
            super().__init__(tag, parent)
            self.setCursor(Qt.PointingHandCursor)
            self.setToolTip(f"点击将 {tag} 插入到当前模板输出路径中")
            self.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            self.setFocusPolicy(Qt.TabFocus)
            self.setAccessibleName(f"变量：{tag}")
            self.apply_theme(THEME.tokens)

        def apply_theme(self, tokens: ColorTokens):
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {tokens.button_bg};
                    border: 1px solid {tokens.border_subtle};
                    border-radius: 11px;
                    padding: 3px 9px;
                    font-size: 11px;
                    color: {tokens.text_primary};
                    text-align: center;
                }}
                QPushButton:hover {{
                    background-color: {tokens.button_hover};
                    border-color: {tokens.accent};
                    color: {tokens.accent};
                }}
                QPushButton:pressed {{
                    background-color: {tokens.button_pressed};
                }}
                QPushButton:focus {{
                    border: 1.5px solid {tokens.accent};
                }}
            """)


    class VariableInsertPanel(QFrame):
        """In-page collapsible panel showing categorized variable capsules with auto-scaling height."""

        MIN_HEIGHT = 65
        MAX_HEIGHT = 280

        variableSelected = Signal(str)

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self.setAccessibleName("输出规则变量选择器")
            self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)

            self._system_vars: list[str] = []
            self._table_vars: list[str] = []
            self._sys_buttons: list[VariableCapsuleButton] = []
            self._tbl_buttons: list[VariableCapsuleButton] = []

            self._build_ui()
            self.apply_theme(THEME.tokens)
            THEME.add_listener(self.apply_theme)
            self.setFixedHeight(self.MIN_HEIGHT)

        def _build_ui(self):
            root_layout = QVBoxLayout(self)
            root_layout.setContentsMargins(12, 8, 12, 8)
            root_layout.setSpacing(6)

            # Top action bar
            top_bar = QHBoxLayout()
            top_bar.setSpacing(8)

            self.lbl_target = QLabel("📌 插入变量到：—")
            top_bar.addWidget(self.lbl_target, 0)

            top_bar.addStretch(1)

            # Search input
            self.edt_search = QLineEdit()
            self.edt_search.setPlaceholderText("🔍 搜索变量...")
            self.edt_search.setFixedWidth(160)
            self.edt_search.setAccessibleName("搜索变量输入框")
            self.edt_search.textChanged.connect(self._on_search_changed)
            top_bar.addWidget(self.edt_search, 0)

            root_layout.addLayout(top_bar)

            # Scroll Area for capsules
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFrameShape(QScrollArea.NoFrame)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            prevent_wheel_propagation(self.scroll_area)

            self.content_widget = QWidget()
            self.content_layout = QVBoxLayout(self.content_widget)
            self.content_layout.setContentsMargins(0, 0, 0, 0)
            self.content_layout.setSpacing(6)

            # Section 1: System Variables
            self.lbl_sys_header = QLabel("🏷️ 系统变量：")
            self.content_layout.addWidget(self.lbl_sys_header)

            self.sys_capsules_widget = QWidget()
            self.sys_capsules_layout = FlowLayout(self.sys_capsules_widget)
            self.content_layout.addWidget(self.sys_capsules_widget)

            # Section 2: Table Columns
            self.lbl_tbl_header = QLabel("📊 数据表格列：")
            self.content_layout.addWidget(self.lbl_tbl_header)

            self.tbl_capsules_widget = QWidget()
            self.tbl_capsules_layout = FlowLayout(self.tbl_capsules_widget)
            self.content_layout.addWidget(self.tbl_capsules_widget)

            # Empty search result hint
            self.lbl_empty_search = QLabel("未找到匹配的变量。")
            self.lbl_empty_search.hide()
            self.content_layout.addWidget(self.lbl_empty_search)

            self.content_layout.addStretch()
            self.scroll_area.setWidget(self.content_widget)
            root_layout.addWidget(self.scroll_area, 1)

        def apply_theme(self, tokens: ColorTokens):
            """Refresh local token styles on theme change."""
            self.setStyleSheet(f"""
                VariableInsertPanel {{
                    background-color: {tokens.card_bg};
                    border: 1px solid {tokens.border_subtle};
                    border-radius: 8px;
                }}
            """)
            self.lbl_target.setStyleSheet(f"font-weight: 600; font-size: 12px; color: {tokens.text_primary};")
            self.edt_search.setStyleSheet(f"""
                QLineEdit {{
                    border: 1px solid {tokens.entry_border};
                    border-radius: 12px;
                    background-color: {tokens.entry_bg};
                    color: {tokens.text_primary};
                    padding: 2px 8px;
                    font-size: 11px;
                }}
                QLineEdit:focus {{
                    border: 1.5px solid {tokens.accent};
                }}
            """)
            self.lbl_sys_header.setStyleSheet(f"font-size: 11px; color: {tokens.text_secondary}; font-weight: 500;")
            self.lbl_tbl_header.setStyleSheet(f"font-size: 11px; color: {tokens.text_secondary}; font-weight: 500; margin-top: 4px;")
            self.lbl_empty_search.setStyleSheet(f"color: {tokens.text_secondary}; font-size: 11px; font-style: italic; padding: 4px;")
            for btn in self._sys_buttons + self._tbl_buttons:
                btn.apply_theme(tokens)

        def adjust_height_to_content(self):
            """Dynamically adjust panel height based on variable content height."""
            w = max(self.width(), 400)
            avail_w = max(w - 36, 200)

            top_bar_h = max(self.edt_search.sizeHint().height(), 26) + 8
            content_h = 0

            if not self.lbl_sys_header.isHidden():
                content_h += self.lbl_sys_header.sizeHint().height() + 4
            if not self.sys_capsules_widget.isHidden():
                content_h += max(self.sys_capsules_layout.heightForWidth(avail_w), 24) + 6

            if not self.lbl_tbl_header.isHidden():
                content_h += self.lbl_tbl_header.sizeHint().height() + 4
            if not self.tbl_capsules_widget.isHidden():
                content_h += max(self.tbl_capsules_layout.heightForWidth(avail_w), 24) + 6

            if not self.lbl_empty_search.isHidden():
                content_h += self.lbl_empty_search.sizeHint().height() + 6

            margins = self.layout().contentsMargins()
            total_h = top_bar_h + content_h + margins.top() + margins.bottom() + 16

            target_h = max(self.MIN_HEIGHT, min(self.MAX_HEIGHT, total_h))
            self.setFixedHeight(target_h)

        def resizeEvent(self, event):
            super().resizeEvent(event)
            if event.oldSize().width() != event.size().width():
                self.adjust_height_to_content()

        def showEvent(self, event):
            super().showEvent(event)
            self.adjust_height_to_content()

        def set_target_name(self, template_name: str):
            """Update the header to reflect which template is being edited."""
            self.lbl_target.setText(f"📌 插入变量到：【{template_name}】")

        def set_catalog(
            self,
            system_vars: list[str],
            table_vars: list[str],
        ):
            """Populate the capsule buttons from system variables and data table headers."""
            self._system_vars = list(system_vars)
            self._table_vars = list(table_vars)

            # Clear existing system buttons
            self._clear_layout(self.sys_capsules_layout)
            self._sys_buttons.clear()

            for tag in self._system_vars:
                btn = VariableCapsuleButton(tag, self)
                btn.clicked.connect(lambda _, t=tag: self._on_capsule_clicked(t))
                self.sys_capsules_layout.addWidget(btn)
                self._sys_buttons.append(btn)

            # Clear existing table buttons
            self._clear_layout(self.tbl_capsules_layout)
            self._tbl_buttons.clear()

            if self._table_vars:
                self.lbl_tbl_header.show()
                self.tbl_capsules_widget.show()
                for col in self._table_vars:
                    tag = f"{{{{{col}}}}}"
                    btn = VariableCapsuleButton(tag, self)
                    btn.clicked.connect(lambda _, t=tag: self._on_capsule_clicked(t))
                    self.tbl_capsules_layout.addWidget(btn)
                    self._tbl_buttons.append(btn)
            else:
                self.lbl_tbl_header.hide()
                self.tbl_capsules_widget.hide()

            self.edt_search.clear()
            self._apply_filter("")

        def _clear_layout(self, layout):
            while layout.count() > 0:
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
                elif child.layout():
                    self._clear_layout(child.layout())

        def _on_capsule_clicked(self, tag: str):
            self.variableSelected.emit(tag)

        def _on_search_changed(self, text: str):
            self._apply_filter(text.strip().lower())

        def _apply_filter(self, query: str):
            sys_visible_count = 0
            for btn in self._sys_buttons:
                match = (not query) or (query in btn.tag.lower())
                btn.setVisible(match)
                if match:
                    sys_visible_count += 1

            tbl_visible_count = 0
            for btn in self._tbl_buttons:
                match = (not query) or (query in btn.tag.lower())
                btn.setVisible(match)
                if match:
                    tbl_visible_count += 1

            self.lbl_sys_header.setVisible(sys_visible_count > 0)
            self.sys_capsules_widget.setVisible(sys_visible_count > 0)
            self.lbl_tbl_header.setVisible(tbl_visible_count > 0 and bool(self._table_vars))
            self.tbl_capsules_widget.setVisible(tbl_visible_count > 0 and bool(self._table_vars))

            has_any = (sys_visible_count + tbl_visible_count) > 0
            self.lbl_empty_search.setVisible(not has_any and bool(query))
            self.sys_capsules_layout.invalidate()
            self.tbl_capsules_layout.invalidate()
            self.content_widget.updateGeometry()
            self.adjust_height_to_content()

        def reset_search(self):
            self.edt_search.clear()


    def create_variable_icon() -> QIcon:
        """Create a clean, crisp '{ }' variable badge icon for QLineEdit trailing action."""
        pixmap = QPixmap(18, 18)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QColor(THEME.tokens.accent))
        font = painter.font()
        font.setPointSize(9)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "{ }")
        painter.end()
        return QIcon(pixmap)

else:

    class VariableInsertPanel:  # type: ignore
        pass

    class VariableCapsuleButton:  # type: ignore
        pass

    def create_variable_icon():  # type: ignore
        return None
