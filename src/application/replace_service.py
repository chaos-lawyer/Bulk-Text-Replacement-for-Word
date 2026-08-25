"""Application service orchestrating text replacement and preview jobs."""

from __future__ import annotations

import re
from typing import Callable, Optional

from application.task_models import CancellationToken, ServiceResult, TaskProgress, TaskState
from core.format_policy import validate_batch_file_formats
from core.models import BatchProcessResult
from core.replacer_core import (
    compile_search_pattern,
    perform_com_preview,
    perform_com_replace,
    perform_standard_preview,
    perform_standard_replace,
    scan_hyperlinks,
)
from platform_adapter.capabilities import CAPABILITIES


class ReplaceService:
    """Encapsulates validation and execution of document search and replace tasks."""

    @staticmethod
    def validate_inputs(
        file_paths: list[str],
        search_text: str,
        mode: str = "fast",
        use_regex: bool = False,
        case_sensitive: bool = False,
        whole_word: bool = False,
    ) -> str | None:
        if not file_paths:
            return "请先添加至少一个 Word 文档。"
        if not search_text:
            return "请输入要查找的内容。"

        # Validate file formats and mode constraints
        ok, format_err = validate_batch_file_formats(file_paths, mode=mode, operation="replace")
        if not ok:
            return format_err

        # Pre-compile regex to detect invalid patterns early
        if use_regex:
            try:
                compile_search_pattern(search_text, case_sensitive=case_sensitive, use_regex=True, whole_word=whole_word)
            except re.error as exc:
                return f"正则表达式语法错误：{exc}"

        return None

    @staticmethod
    def execute_preview(
        file_paths: list[str],
        search_text: str,
        mode: str = "fast",
        case_sensitive: bool = False,
        use_regex: bool = False,
        whole_word: bool = False,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[BatchProcessResult]:
        err = ReplaceService.validate_inputs(
            file_paths,
            search_text,
            mode=mode,
            use_regex=use_regex,
            case_sensitive=case_sensitive,
            whole_word=whole_word,
        )
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, filename: str):
            if progress_cb:
                progress_cb(TaskProgress.calculate(current, total, f"正在分析：{filename}"))

        try:
            if mode == "full":
                if not CAPABILITIES.has_word_com:
                    return ServiceResult(
                        success=False,
                        error="完整模式需要 Windows、Microsoft Word 和 pywin32 支持。",
                        state=TaskState.FAILED,
                    )
                res = perform_com_preview(
                    file_paths=file_paths,
                    search_text=search_text,
                    case_sensitive=case_sensitive,
                    whole_word=whole_word,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )
            else:
                res = perform_standard_preview(
                    file_paths=file_paths,
                    search_text=search_text,
                    case_sensitive=case_sensitive,
                    use_regex=use_regex,
                    whole_word=whole_word,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=res, error="任务已取消", state=TaskState.CANCELLED)

            if res.total_files > 0 and res.files_processed == 0 and res.errors:
                return ServiceResult(success=False, data=res, error="\n".join(res.errors), state=TaskState.FAILED)

            state = TaskState.WARNING if res.errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=res, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def execute_replace(
        file_paths: list[str],
        search_text: str,
        replace_text: str,
        mode: str = "fast",
        case_sensitive: bool = False,
        use_regex: bool = False,
        whole_word: bool = False,
        create_backup: bool = True,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[BatchProcessResult]:
        err = ReplaceService.validate_inputs(
            file_paths,
            search_text,
            mode=mode,
            use_regex=use_regex,
            case_sensitive=case_sensitive,
            whole_word=whole_word,
        )
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, filename: str):
            if progress_cb:
                progress_cb(TaskProgress.calculate(current, total, f"正在替换：{filename}"))

        try:
            if mode == "full":
                if not CAPABILITIES.has_word_com:
                    return ServiceResult(
                        success=False,
                        error="完整模式需要 Windows、Microsoft Word 和 pywin32 支持。",
                        state=TaskState.FAILED,
                    )
                res = perform_com_replace(
                    file_paths=file_paths,
                    search_text=search_text,
                    replace_text=replace_text,
                    case_sensitive=case_sensitive,
                    whole_word=whole_word,
                    create_backup=create_backup,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )
            else:
                res = perform_standard_replace(
                    file_paths=file_paths,
                    search_text=search_text,
                    replace_text=replace_text,
                    case_sensitive=case_sensitive,
                    use_regex=use_regex,
                    whole_word=whole_word,
                    create_backup=create_backup,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=res, error="任务已取消", state=TaskState.CANCELLED)

            if res.total_files > 0 and res.successful_files == 0 and res.errors:
                return ServiceResult(success=False, data=res, error="\n".join(res.errors), state=TaskState.FAILED)

            state = TaskState.WARNING if res.errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=res, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def scan_links(
        file_paths: list[str],
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[list[dict]]:
        if not file_paths:
            return ServiceResult(success=False, error="请先添加 Word 文档。", state=TaskState.FAILED)

        try:
            if progress_cb:
                progress_cb(TaskProgress.calculate(1, len(file_paths), "正在扫描超链接..."))

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            links = scan_hyperlinks(file_paths)
            errors = [f"{item['filename']}: {item['error']}" for item in links if item.get("error")]

            if len(errors) == len(file_paths) and len(file_paths) > 0:
                return ServiceResult(success=False, data=links, error="\n".join(errors), state=TaskState.FAILED)

            state = TaskState.WARNING if errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=links, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)
