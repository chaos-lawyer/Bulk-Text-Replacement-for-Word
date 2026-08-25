"""Application service orchestrating multi-document differentiated variable scanning and replacement."""

from __future__ import annotations

import os
from typing import Callable, Optional

from application.task_models import CancellationToken, ServiceResult, TaskProgress, TaskState
from core.format_policy import validate_batch_file_formats
from core.models import MultiDocBatchResult, MultiDocItem
from core.multi_doc_replacer import (
    execute_multi_doc_preview,
    execute_multi_doc_replace,
    scan_documents_variables,
)


class MultiDocService:
    """Encapsulates validation, scanning, preview generation, and batch replacement for multiple documents."""

    @staticmethod
    def validate_inputs(
        items: list[MultiDocItem],
        in_place: bool = True,
        output_folder: str | None = None,
        use_com: bool = False,
    ) -> str | None:
        if not items:
            return "请先添加至少一个待处理的 Word 文档。"

        file_paths = [item.file_path for item in items]
        ok, format_err = validate_batch_file_formats(
            file_paths, mode="full" if use_com else "fast", operation="multi_doc"
        )
        if not ok:
            return format_err

        if not in_place:
            if not output_folder or not output_folder.strip():
                return "请选择生成文档的输出保存目录。"

        return None

    @staticmethod
    def scan_documents(
        file_paths: list[str],
        use_com: bool = False,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[tuple[dict[str, list[str]], list[str], dict[str, str]]]:
        if not file_paths:
            return ServiceResult(success=False, error="请先添加待扫描的 Word 文档。", state=TaskState.FAILED)

        ok, format_err = validate_batch_file_formats(
            file_paths, mode="full" if use_com else "fast", operation="multi_doc"
        )
        if not ok:
            return ServiceResult(success=False, error=format_err, state=TaskState.FAILED)

        if cancel_token and cancel_token.is_cancelled:
            return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

        def _inner_progress(current: int, total: int, filename: str):
            if progress_cb:
                progress_cb(TaskProgress.calculate(current, total, f"正在扫描变量：{filename}"))

        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        try:
            doc_vars, all_vars, doc_errors = scan_documents_variables(
                file_paths=file_paths,
                use_com=use_com,
                progress=_inner_progress,
                is_cancelled=is_cancelled_fn,
            )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            if doc_errors:
                if len(doc_errors) == len(file_paths) and len(file_paths) > 0:
                    err_summary = "\n".join(f"{os.path.basename(p)}: {e}" for p, e in doc_errors.items())
                    return ServiceResult(
                        success=False,
                        data=(doc_vars, all_vars, doc_errors),
                        error=f"全部文档扫描失败：\n{err_summary}",
                        state=TaskState.FAILED,
                    )
                state = TaskState.WARNING
            else:
                state = TaskState.SUCCESS

            return ServiceResult(
                success=True,
                data=(doc_vars, all_vars, doc_errors),
                state=state,
            )
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def generate_preview(
        items: list[MultiDocItem],
        all_variables: list[str],
        output_folder: str | None = None,
        max_items: int = 10,
    ) -> ServiceResult[list[dict]]:
        try:
            snippets = execute_multi_doc_preview(
                items=items,
                all_variables=all_variables,
                output_folder=output_folder,
                max_items=max_items,
            )
            return ServiceResult(success=True, data=snippets, state=TaskState.SUCCESS)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def execute_batch_replace(
        items: list[MultiDocItem],
        in_place: bool = True,
        output_folder: str | None = None,
        create_backup: bool = True,
        use_com: bool = False,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[MultiDocBatchResult]:
        err = MultiDocService.validate_inputs(items, in_place=in_place, output_folder=output_folder, use_com=use_com)
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        target_output = None if in_place else output_folder
        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, doc_detail: dict):
            if progress_cb:
                prog = TaskProgress.calculate(
                    current, total, f"正在处理：{doc_detail['filename']}", data=doc_detail
                )
                progress_cb(prog)

        try:
            batch_res = execute_multi_doc_replace(
                items=items,
                output_folder=target_output,
                create_backup=create_backup,
                use_com=use_com,
                progress=_inner_progress,
                is_cancelled=is_cancelled_fn,
            )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(
                    success=False, data=batch_res, error="任务已取消", state=TaskState.CANCELLED
                )

            if batch_res.total_docs > 0 and batch_res.success_docs == 0 and batch_res.errors:
                return ServiceResult(
                    success=False, data=batch_res, error="\n".join(batch_res.errors), state=TaskState.FAILED
                )

            state = TaskState.WARNING if batch_res.errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=batch_res, state=state)

        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)
