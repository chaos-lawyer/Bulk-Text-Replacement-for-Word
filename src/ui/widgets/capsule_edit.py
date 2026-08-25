"""Capsule / Chip file tag container widget with drag-and-drop and removal capabilities."""

from __future__ import annotations

import os
from pathlib import Path
from ui.theme.theme_manager import THEME

try:
    from PySide6.QtCore import Qt, Signal
    from PySide6.QtGui import QDragEnterEvent, QDropEvent
    from PySide6.QtWidgets import (
        QFrame,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QSizePolicy,
        QToolButton,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FileCapsuleItem(QFrame):
        """A single rounded pill/capsule chip representing a file item with a close 'x' button."""

        removed = Signal(str)  # Emits file_path

        def __init__(self, file_path: str, display_name: str = "", parent: QWidget | None = None):
            super().__init__(parent)
            self.file_path = file_path
            self.display_name = display_name or Path(file_path).name

            self.setToolTip(f"文件：{self.display_name}\n完整路径：{self.file_path}")
            self.setCursor(Qt.PointingHandCursor)
            self._setup_ui()
            self._apply_style()

        def _setup_ui(self):
            layout = QHBoxLayout(self)
            layout.setContentsMargins(8, 2, 6, 2)
            layout.setSpacing(6)

            # Icon / Label
            self.lbl_icon = QLabel("📄")
            self.lbl_icon.setStyleSheet("font-size: 11px;")
            layout.addWidget(self.lbl_icon)

            self.lbl_text = QLabel(self.display_name)
            self.lbl_text.setStyleSheet("font-size: 12px; font-weight: 500;")
            layout.addWidget(self.lbl_text)

            # Close button
            self.btn_close = QToolButton(self)
            self.btn_close.setText("✕")
            self.btn_close.setCursor(Qt.PointingHandCursor)
            self.btn_close.setFixedSize(16, 16)
            self.btn_close.setToolTip("移除此模板")
            self.btn_close.clicked.connect(lambda: self.removed.emit(self.file_path))
            layout.addWidget(self.btn_close)

        def _apply_style(self, tokens: ColorTokens | None = None):
            t = tokens or THEME.tokens
            self.setStyleSheet(f"""
                QFrame {{
                    background-color: {t.badge_bg};
                    border: 1px solid {t.border};
                    border-radius: 12px;
                }}
                QFrame:hover {{
                    background-color: {t.button_hover};
                    border: 1px solid {t.border_strong};
                }}
                QLabel {{
                    color: {t.text_primary};
                    background: transparent;
                    border: none;
                }}
                QToolButton {{
                    color: {t.text_secondary};
                    background: transparent;
                    border: none;
                    border-radius: 8px;
                    font-size: 10px;
                    font-weight: bold;
                    padding: 0px;
                }}
                QToolButton:hover {{
                    color: {t.error};
                    background-color: {t.button_pressed};
                }}
            """)


    class FileCapsuleEdit(QFrame):
        """An input-box container widget that renders selected files as text capsules and accepts drag & drop."""

        filesDropped = Signal(list)  # list[str]
        fileRemoved = Signal(str)    # file_path
        clicked = Signal()

        def __init__(
            self,
            placeholder_text: str = "点击“添加模板…”或将 Word 模板文件直接拖放到此处...",
            supported_extensions: tuple[str, ...] = (".docx", ".docm", ".doc"),
            parent: QWidget | None = None,
        ):
            super().__init__(parent)
            self.placeholder_text = placeholder_text
            self.supported_extensions = tuple(ext.lower() for ext in supported_extensions)
            self._file_paths: list[str] = []
            self._is_drag_hover = False

            self.setAcceptDrops(True)
            self.setMinimumHeight(32)
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            self.setCursor(Qt.PointingHandCursor)

            self._build_ui()
            self._update_appearance()
            THEME.add_listener(self._on_theme_changed)

        def _on_theme_changed(self, tokens):
            self._update_appearance(tokens)
            for i in range(self.capsule_layout.count()):
                item = self.capsule_layout.itemAt(i)
                if item and item.widget() and isinstance(item.widget(), FileCapsuleItem):
                    item.widget()._apply_style(tokens)

        def _build_ui(self):
            self.main_layout = QHBoxLayout(self)
            self.main_layout.setContentsMargins(6, 4, 6, 4)
            self.main_layout.setSpacing(6)

            # Placeholder label
            self.lbl_placeholder = QLabel(self.placeholder_text, self)
            self.lbl_placeholder.setStyleSheet(f"color: {THEME.tokens.text_placeholder}; font-size: 12px;")
            self.lbl_placeholder.setCursor(Qt.PointingHandCursor)
            self.main_layout.addWidget(self.lbl_placeholder)

            # Capsule items container
            self.capsule_container = QWidget(self)
            self.capsule_layout = QHBoxLayout(self.capsule_container)
            self.capsule_layout.setContentsMargins(0, 0, 0, 0)
            self.capsule_layout.setSpacing(6)
            self.main_layout.addWidget(self.capsule_container)
            self.capsule_container.hide()

            self.main_layout.addStretch()

        def set_files(self, items: list[dict | tuple[str, str]] | list[str]):
            """Set file list and rebuild capsule badges."""
            self._file_paths.clear()

            # Clear existing capsule widgets
            while self.capsule_layout.count() > 0:
                child = self.capsule_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

            for it in items:
                if isinstance(it, tuple):
                    path, name = it[0], it[1]
                elif isinstance(it, dict):
                    path, name = it.get("path", ""), it.get("name", "")
                else:
                    path, name = str(it), Path(str(it)).name

                if path:
                    self._file_paths.append(path)
                    capsule = FileCapsuleItem(file_path=path, display_name=name, parent=self.capsule_container)
                    capsule.removed.connect(self._on_capsule_removed)
                    self.capsule_layout.addWidget(capsule)

            has_files = len(self._file_paths) > 0
            self.lbl_placeholder.setVisible(not has_files)
            self.capsule_container.setVisible(has_files)

        def _on_capsule_removed(self, path: str):
            self.fileRemoved.emit(path)

        def _update_appearance(self, tokens: ColorTokens | None = None):
            t = tokens or THEME.tokens
            border_color = t.accent if self._is_drag_hover else t.entry_border
            bg_color = t.button_hover if self._is_drag_hover else t.entry_bg
            border_width = "1.5px" if self._is_drag_hover else "1px"

            self.setStyleSheet(f"""
                QFrame {{
                    background-color: {bg_color};
                    border: {border_width} solid {border_color};
                    border-radius: 6px;
                }}
            """)
            self.lbl_placeholder.setStyleSheet(f"color: {t.text_placeholder}; font-size: 12px; background: transparent; border: none;")

        def mousePressEvent(self, event):
            super().mousePressEvent(event)
            self.clicked.emit()

        def dragEnterEvent(self, event: QDragEnterEvent):
            if event.mimeData().hasUrls():
                urls = event.mimeData().urls()
                valid = any(
                    Path(url.toLocalFile()).suffix.lower() in self.supported_extensions
                    for url in urls
                    if url.isLocalFile()
                )
                if valid:
                    event.acceptProposedAction()
                    self._is_drag_hover = True
                    self._update_appearance()
                    return
            event.ignore()

        def dragMoveEvent(self, event):
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
            else:
                event.ignore()

        def dragLeaveEvent(self, event):
            self._is_drag_hover = False
            self._update_appearance()
            super().dragLeaveEvent(event)

        def dropEvent(self, event: QDropEvent):
            self._is_drag_hover = False
            self._update_appearance()

            if event.mimeData().hasUrls():
                paths = []
                for url in event.mimeData().urls():
                    if url.isLocalFile():
                        fpath = url.toLocalFile()
                        if Path(fpath).suffix.lower() in self.supported_extensions:
                            paths.append(os.path.abspath(fpath))
                if paths:
                    event.acceptProposedAction()
                    self.filesDropped.emit(paths)
                    return
            event.ignore()


    class DropLineEdit(QLineEdit):
        """QLineEdit supporting file drag-and-drop with highlight feedback."""

        fileDropped = Signal(str)

        def __init__(
            self,
            supported_extensions: tuple[str, ...] = (".xlsx", ".csv", ".xlsm", ".xltx", ".xltm"),
            parent: QWidget | None = None,
        ):
            super().__init__(parent)
            self.supported_extensions = tuple(ext.lower() for ext in supported_extensions)
            self._is_drag_hover = False
            self.setAcceptDrops(True)

        def dragEnterEvent(self, event: QDragEnterEvent):
            if event.mimeData().hasUrls():
                for url in event.mimeData().urls():
                    if url.isLocalFile():
                        fpath = url.toLocalFile()
                        if not self.supported_extensions or Path(fpath).suffix.lower() in self.supported_extensions:
                            event.acceptProposedAction()
                            self._is_drag_hover = True
                            self._apply_drag_style()
                            return
            super().dragEnterEvent(event)

        def dragMoveEvent(self, event):
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
            else:
                super().dragMoveEvent(event)

        def dragLeaveEvent(self, event):
            self._is_drag_hover = False
            self._clear_drag_style()
            super().dragLeaveEvent(event)

        def dropEvent(self, event: QDropEvent):
            self._is_drag_hover = False
            self._clear_drag_style()

            if event.mimeData().hasUrls():
                for url in event.mimeData().urls():
                    if url.isLocalFile():
                        fpath = url.toLocalFile()
                        if not self.supported_extensions or Path(fpath).suffix.lower() in self.supported_extensions:
                            event.acceptProposedAction()
                            abs_p = os.path.abspath(fpath)
                            self.setText(abs_p)
                            self.fileDropped.emit(abs_p)
                            return
            super().dropEvent(event)

        def _apply_drag_style(self):
            tokens = THEME.tokens
            self.setStyleSheet(f"""
                QLineEdit {{
                    border: 1.5px solid {tokens.accent};
                    background-color: {tokens.button_hover};
                }}
            """)

        def _clear_drag_style(self):
            self.setStyleSheet("")

else:

    class FileCapsuleItem:  # type: ignore
        pass

    class FileCapsuleEdit:  # type: ignore
        pass

    class DropLineEdit:  # type: ignore
        pass
