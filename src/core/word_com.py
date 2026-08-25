"""Word COM Automation Session context manager and story range traversal helpers."""

from __future__ import annotations

from typing import Any, Callable

from platform_adapter.capabilities import CAPABILITIES


class WordAutomationSession:
    """Manages Microsoft Word COM lifecycle with per-thread apartment initialization and safe teardown.

    Guarantees:
    1. Calls pythoncom.CoInitialize() on enter and pythoncom.CoUninitialize() on exit.
    2. Uses DispatchEx to launch an isolated Word process instance.
    3. Hides window and disables interactive alerts.
    4. Quits Word and releases COM proxies on exit, even if exceptions occur.
    """

    def __init__(self, visible: bool = False, display_alerts: bool = False):
        self.visible = visible
        self.display_alerts = display_alerts
        self.app: Any = None
        self._initialized_com = False

    def __enter__(self) -> Any:
        if not CAPABILITIES.is_windows:
            raise RuntimeError("Word COM 自动化仅支持 Windows 平台。")

        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError("需要安装 pywin32 模块以支持 Word COM 完整模式。") from exc

        try:
            pythoncom.CoInitialize()
            self._initialized_com = True
        except Exception:
            pass

        try:
            # DispatchEx ensures a new, isolated WINWORD.EXE process is spawned
            self.app = win32com.client.DispatchEx("Word.Application")
            self.app.Visible = self.visible
            self.app.DisplayAlerts = 0 if not self.display_alerts else -1
            try:
                self.app.ScreenUpdating = False
            except Exception:
                pass
            return self.app
        except Exception as exc:
            self._cleanup()
            raise RuntimeError(f"无法启动 Microsoft Word COM 自动化服务，请确认已安装 Microsoft Word：{exc}") from exc

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup()

    def _cleanup(self):
        if self.app is not None:
            try:
                self.app.ScreenUpdating = True
            except Exception:
                pass
            try:
                self.app.Quit(0)  # 0 = wdDoNotSaveChanges
            except Exception:
                pass
            self.app = None

        if self._initialized_com:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._initialized_com = False


def traverse_story_ranges(
    document: Any,
    operation_fn: Callable[[Any], int],
) -> tuple[int, list[str]]:
    """Traverse all story ranges in a Word document (Body, Headers, Footers, Shapes, TextBoxes, Footnotes, Endnotes).

    Returns (total_count, warnings_list).
    """
    total = 0
    warnings = []

    for first_story in document.StoryRanges:
        story = first_story
        while story is not None:
            try:
                total += operation_fn(story)
            except Exception as exc:
                warnings.append(f"区域处理异常 ({getattr(story, 'StoryType', '未知')}): {exc}")

            # Also check text frame shapes inside this story range if any
            try:
                if hasattr(story, "ShapeRange") and story.ShapeRange.Count > 0:
                    for shape in story.ShapeRange:
                        try:
                            if shape.TextFrame.HasText:
                                total += operation_fn(shape.TextFrame.TextRange)
                        except Exception:
                            pass
            except Exception:
                pass

            try:
                story = story.NextStoryRange
            except Exception:
                story = None

    return total, warnings
