"""Application service orchestrating Word template(s) + Excel/CSV batch generation."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Callable, Mapping, Optional

from application.task_models import CancellationToken, ServiceResult, TaskProgress, TaskState
from core.format_policy import validate_batch_file_formats
from core.models import ExcelData, MergePlanResult, MergeResult, TemplateMergeItem
from core.template_merge import (
    build_field_mapping,
    create_merge_plan,
    extract_template_fields,
    extract_template_fields_com,
    generate_multi_template_batch,
    load_table_data,
    resolve_empty_field,
)
from platform_adapter.capabilities import CAPABILITIES


class MergeService:
    """Encapsulates validation, scanning, planning, and execution for multi-template batch merge tasks."""

    @staticmethod
    def validate_sources(
        template_paths: list[str] | str,
        excel_path: str,
        use_com: bool = False,
    ) -> str | None:
        if isinstance(template_paths, str):
            paths = [template_paths] if template_paths.strip() else []
        else:
            paths = template_paths

        if not paths:
            return "请至少添加一个 Word 模板文件（.docx/.docm/.doc）。"

        ok, format_err = validate_batch_file_formats(
            paths, mode="full" if use_com else "fast", operation="merge"
        )
        if not ok:
            return format_err

        if not excel_path or not os.path.isfile(excel_path):
            return "请选择存在的表格数据文件（.xlsx/.csv/.xlsm）。"

        table_suffix = Path(excel_path).suffix.lower()
        if table_suffix not in {".xlsx", ".xlsm", ".csv", ".xltx", ".xltm"}:
            return "数据源文件必须为 .xlsx、.xlsm 或 .csv 格式表格。"

        return None

    @staticmethod
    def scan_templates_and_excel(
        templates: list[TemplateMergeItem],
        excel_path: str,
        use_com: bool = False,
        sheet_name: Optional[str] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[tuple[list[TemplateMergeItem], list[str], dict[str, list[str]], ExcelData, dict[str, str]]]:
        """Scan multiple Word templates, extract field union and template mappings, and load Excel data."""
        paths = [t.file_path for t in templates if t.enabled]
        if not paths:
            paths = [t.file_path for t in templates]

        err = MergeService.validate_sources(paths, excel_path, use_com=use_com)
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        if cancel_token and cancel_token.is_cancelled:
            return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

        all_fields_ordered: list[str] = []
        all_fields_seen: set[str] = set()
        field_to_templates: dict[str, list[str]] = {}
        total_tmpls = len(templates)
        scan_errors: list[str] = []

        try:
            for idx, tmpl in enumerate(templates, start=1):
                if cancel_token and cancel_token.is_cancelled:
                    return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

                if progress_cb:
                    progress_cb(TaskProgress.calculate(idx, total_tmpls + 1, f"正在扫描模板：{tmpl.display_name}"))

                if not tmpl.enabled:
                    continue

                suffix = Path(tmpl.file_path).suffix.lower()
                try:
                    if use_com or suffix == ".doc":
                        if not CAPABILITIES.has_word_com:
                            raise RuntimeError("Word COM 完整模式需要 Windows 与 Microsoft Word 支持")
                        fields = extract_template_fields_com(tmpl.file_path)
                    else:
                        fields = extract_template_fields(tmpl.file_path)
                    tmpl.fields = fields
                    tmpl.status = f"已检测到 {len(fields)} 个变量"
                    tmpl.error = ""
                except Exception as exc:
                    tmpl.fields = []
                    tmpl.status = "扫描失败"
                    tmpl.error = str(exc)
                    scan_errors.append(f"{tmpl.display_name}: {exc}")
                    continue

                for f in fields:
                    if f not in all_fields_seen:
                        all_fields_seen.add(f)
                        all_fields_ordered.append(f)
                    field_to_templates.setdefault(f, []).append(tmpl.display_name)

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            if progress_cb:
                progress_cb(TaskProgress.calculate(total_tmpls + 1, total_tmpls + 1, "正在读取表格数据..."))

            excel_data = load_table_data(excel_path, sheet_name=sheet_name)
            mapping = build_field_mapping(all_fields_ordered, excel_data.headers)

            if not all_fields_ordered and not scan_errors:
                return ServiceResult(
                    success=False,
                    error="所有选中的 Word 模板中均未检测到 {{字段名}} 格式的变量。",
                    state=TaskState.FAILED,
                )

            state = TaskState.WARNING if scan_errors else TaskState.SUCCESS
            return ServiceResult(
                success=True,
                data=(templates, all_fields_ordered, field_to_templates, excel_data, mapping),
                state=state,
            )
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def scan_template_and_excel(
        template_path: str,
        excel_path: str,
        use_com: bool = False,
        sheet_name: Optional[str] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[tuple[list[str], ExcelData, dict[str, str]]]:
        """Single-template scan wrapper for backward compatibility."""
        tmpl = TemplateMergeItem(
            template_id="tmpl_1",
            file_path=template_path,
            display_name=Path(template_path).stem,
            enabled=True,
        )
        res = MergeService.scan_templates_and_excel(
            templates=[tmpl],
            excel_path=excel_path,
            use_com=use_com,
            sheet_name=sheet_name,
            progress_cb=progress_cb,
            cancel_token=cancel_token,
        )
        if not res.success:
            return ServiceResult(success=False, error=res.error, state=res.state)

        _tmpls, all_fields, _field_tmpls, excel_data, mapping = res.data
        return ServiceResult(success=True, data=(all_fields, excel_data, mapping), state=res.state)

    @staticmethod
    def generate_preview_snippets(
        templates: list[TemplateMergeItem] | str | None = None,
        excel_data: Optional[ExcelData] = None,
        mapping: Optional[Mapping[str, str]] = None,
        output_folder: str = "",
        filename_rule: str = "",
        relative_folder_rule: str = "",
        default_values: Optional[Mapping[str, str]] = None,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
        max_rows: int = 5,
        template_path: Optional[str] = None,
        output_directory_rule: str = "",
        plan: Optional[MergePlanResult] = None,
    ) -> list[dict]:
        """Generate structured preview snippets for multi-template batch merge jobs."""
        if templates is None and template_path is not None:
            templates = template_path

        if isinstance(templates, str):
            tmpl_items = [
                TemplateMergeItem(
                    template_id="tmpl_1",
                    file_path=templates,
                    display_name=Path(templates).stem,
                    filename_rule=filename_rule or "{{模板名}}-{{数据序号}}",
                    enabled=True,
                )
            ]
        elif isinstance(templates, list):
            tmpl_items = templates
        else:
            tmpl_items = []

        if excel_data is None:
            return []

        mapping = mapping or {}
        out_dir_rule = output_directory_rule or output_folder

        active_plan = plan or create_merge_plan(
            templates=tmpl_items,
            data=excel_data,
            mapping=mapping,
            output_directory_rule=out_dir_rule,
            base_output_folder=output_folder if not out_dir_rule else None,
            relative_folder_rule=relative_folder_rule,
            default_filename_rule=filename_rule,
            default_values=default_values,
            empty_field_behaviors=empty_field_behaviors,
        )
        jobs = active_plan.jobs

        snippets = []
        row_lookup = {
            d_row: r_dict
            for d_row, r_dict in zip(excel_data.excel_rows, excel_data.rows)
        }

        for job in jobs[: max_rows * len(tmpl_items)]:
            row_dict = row_lookup.get(job.excel_row, {})

            field_pairs = []
            for field, header in mapping.items():
                raw_val = row_dict.get(header, "") if header else ""
                should_replace, resolved, behavior = resolve_empty_field(
                    field,
                    raw_val,
                    default_values=default_values,
                    empty_field_behaviors=empty_field_behaviors,
                )
                field_pairs.append(
                    {
                        "field": field,
                        "header": header,
                        "value": resolved if should_replace else f"{{{{{field}}}}}",
                        "is_mapped": bool(header),
                        "is_default": behavior == "custom",
                        "empty_behavior": behavior,
                    }
                )

            snippets.append(
                {
                    "template_name": job.template_name,
                    "excel_row": job.excel_row,
                    "data_index": job.data_index,
                    "relative_folder": job.relative_folder,
                    "filename": job.filename,
                    "destination": job.destination,
                    "collision_reason": job.collision_reason,
                    "fields": field_pairs,
                }
            )

        return snippets

    @staticmethod
    def prepare_merge_preview(
        templates: list[TemplateMergeItem],
        excel_data: ExcelData,
        mapping: Mapping[str, str],
        output_directory_rule: str = "",
        default_values: Optional[Mapping[str, str]] = None,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
        max_rows: int = 5,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[tuple[MergePlanResult, list[dict]]]:
        """Build one reusable plan in the background and render preview rows from it."""
        if cancel_token and cancel_token.is_cancelled:
            return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)
        if progress_cb:
            progress_cb(TaskProgress.calculate(0, 1, "正在规划输出目标..."))

        try:
            plan = create_merge_plan(
                templates=templates,
                data=excel_data,
                mapping=mapping,
                output_directory_rule=output_directory_rule,
                default_values=default_values,
                empty_field_behaviors=empty_field_behaviors,
            )
            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            snippets = MergeService.generate_preview_snippets(
                templates=templates,
                excel_data=excel_data,
                mapping=mapping,
                output_directory_rule=output_directory_rule,
                default_values=default_values,
                empty_field_behaviors=empty_field_behaviors,
                max_rows=max_rows,
                plan=plan,
            )
            if progress_cb:
                progress_cb(TaskProgress.calculate(1, 1, "任务规划已完成"))
            state = TaskState.WARNING if plan.issues else TaskState.SUCCESS
            return ServiceResult(success=True, data=(plan, snippets), state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def execute_multi_template_merge(
        templates: list[TemplateMergeItem],
        excel_data: ExcelData,
        mapping: Mapping[str, str],
        base_output_folder: str = "",
        relative_folder_rule: str = "",
        default_filename_rule: str = "",
        use_com: bool = False,
        default_values: Optional[Mapping[str, str]] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
        replace_empty: bool = True,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
        output_directory_rule: str = "",
        plan: Optional[MergePlanResult] = None,
    ) -> ServiceResult[list[MergeResult]]:
        """Execute multi-template batch merge with progress and cancellation."""
        enabled_tmpls = [t for t in templates if t.enabled]
        if not enabled_tmpls:
            return ServiceResult(success=False, error="请至少启用一个 Word 模板。", state=TaskState.FAILED)

        out_dir_rule = output_directory_rule or base_output_folder
        for tmpl in enabled_tmpls:
            rule = (tmpl.filename_rule or "").strip()
            if not rule and not out_dir_rule:
                return ServiceResult(
                    success=False,
                    error=f"模板【{tmpl.display_name}】未配置输出路径与文件名规则。",
                    state=TaskState.FAILED,
                )

        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, result: MergeResult):
            if progress_cb:
                folder_desc = f"{result.relative_folder}/" if result.relative_folder else ""
                prog = TaskProgress.calculate(
                    current,
                    total,
                    f"[{result.template_name}] 第 {result.excel_row} 行 => {folder_desc}{result.filename}",
                    data=result,
                )
                progress_cb(prog)

        try:
            results = generate_multi_template_batch(
                templates=templates,
                data=excel_data,
                mapping=mapping,
                output_directory_rule=out_dir_rule,
                base_output_folder=base_output_folder if not output_directory_rule else None,
                relative_folder_rule=relative_folder_rule,
                default_filename_rule=default_filename_rule,
                use_com=use_com,
                default_values=default_values,
                empty_field_behaviors=empty_field_behaviors,
                progress=_inner_progress,
                is_cancelled=is_cancelled_fn,
                replace_empty=replace_empty,
                plan=plan,
            )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=results, error="任务已取消", state=TaskState.CANCELLED)

            failed = [r for r in results if not r.success]
            if len(failed) == len(results) and len(results) > 0:
                err_msgs = [f"{r.template_name} (行{r.excel_row}): {r.error}" for r in failed]
                return ServiceResult(success=False, data=results, error="\n".join(err_msgs), state=TaskState.FAILED)

            state = TaskState.WARNING if failed else TaskState.SUCCESS
            return ServiceResult(success=True, data=results, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def execute_batch_merge(
        template_path: str,
        excel_data: ExcelData,
        mapping: Mapping[str, str],
        output_folder: str,
        filename_rule: str,
        use_com: bool = False,
        default_values: Optional[Mapping[str, str]] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
        replace_empty: bool = True,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
    ) -> ServiceResult[list[MergeResult]]:
        """Single-template batch merge wrapper for backward compatibility."""
        tmpl = TemplateMergeItem(
            template_id="tmpl_1",
            file_path=template_path,
            display_name=Path(template_path).stem,
            filename_rule=filename_rule,
            enabled=True,
        )
        return MergeService.execute_multi_template_merge(
            templates=[tmpl],
            excel_data=excel_data,
            mapping=mapping,
            base_output_folder=output_folder,
            relative_folder_rule="",
            default_filename_rule=filename_rule,
            use_com=use_com,
            default_values=default_values,
            progress_cb=progress_cb,
            cancel_token=cancel_token,
            replace_empty=replace_empty,
            empty_field_behaviors=empty_field_behaviors,
        )
