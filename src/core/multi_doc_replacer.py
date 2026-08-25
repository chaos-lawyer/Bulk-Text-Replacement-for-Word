"""Core engine for multi-document differentiated variable scanning and batch replacement.

Decoupled from GUI frameworks for headless execution, automated testing, and thread safety.
"""

from __future__ import annotations

import os
from pathlib import Path
import shutil
from typing import Callable, Optional

from core.file_utils import create_safe_backup
from core.models import MultiDocBatchResult, MultiDocItem
from core.template_merge import (
    _com_replace_document,
    extract_template_fields,
    extract_template_fields_com,
    replace_docx_fields,
)
from core.word_com import WordAutomationSession


def scan_documents_variables(
    file_paths: list[str],
    use_com: bool = False,
    progress: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> tuple[dict[str, list[str]], list[str], dict[str, str]]:
    """Extract template placeholder variables from multiple Word documents.

    Returns:
        tuple[dict[str, list[str]], list[str], dict[str, str]]:
            - Mapping of file_path -> list of detected variable names for that file
            - Deduplicated, order-preserved list of all unique variable names across all files
            - Mapping of file_path -> error message (if scanning failed for that file)
    """
    doc_vars_map: dict[str, list[str]] = {}
    doc_errors: dict[str, str] = {}
    all_vars_ordered: list[str] = []
    all_vars_seen: set[str] = set()

    total = len(file_paths)
    for i, path in enumerate(file_paths, start=1):
        if is_cancelled and is_cancelled():
            break

        filename = os.path.basename(path)
        if progress:
            progress(i, total, filename)

        suffix = Path(path).suffix.lower()
        fields: list[str] = []
        try:
            if not os.path.isfile(path):
                raise FileNotFoundError("文件不存在或无法访问")

            if suffix == ".doc":
                if use_com:
                    fields = extract_template_fields_com(path)
                else:
                    raise ValueError(".doc 格式需启用 Word COM 完整模式进行扫描")
            elif suffix in {".docx", ".docm"}:
                fields = extract_template_fields(path)
            else:
                raise ValueError(f"不支持的格式：{suffix}")
        except Exception as exc:
            fields = []
            doc_errors[path] = str(exc)

        doc_vars_map[path] = fields
        for f in fields:
            if f not in all_vars_seen:
                all_vars_seen.add(f)
                all_vars_ordered.append(f)

    return doc_vars_map, all_vars_ordered, doc_errors


def execute_multi_doc_preview(
    items: list[MultiDocItem],
    all_variables: list[str],
    output_folder: str | None = None,
    max_items: int = 10,
) -> list[dict]:
    """Generate structured preview information for multi-document replacements."""
    preview_data = []
    output_folder_clean = output_folder.strip() if output_folder else ""

    for index, item in enumerate(items[:max_items], start=1):
        if output_folder_clean:
            dest_desc = str(Path(output_folder_clean, item.filename))
        else:
            dest_desc = f"原文件：{item.filename}（将生成安全备份）"

        field_details = []
        for var in all_variables:
            val = item.replacements.get(var, "")
            is_detected = var in item.detected_variables
            field_details.append(
                {
                    "field": var,
                    "value": val,
                    "is_detected": is_detected,
                    "is_filled": bool(val and str(val).strip()),
                }
            )

        preview_data.append(
            {
                "index": index,
                "filename": item.filename,
                "file_path": item.file_path,
                "destination": dest_desc,
                "status": item.status,
                "fields": field_details,
            }
        )

    return preview_data


def execute_multi_doc_replace(
    items: list[MultiDocItem],
    output_folder: str | None = None,
    create_backup: bool = True,
    use_com: bool = False,
    progress: Optional[Callable[[int, int, dict], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> MultiDocBatchResult:
    """Execute differentiated variable replacement across multiple documents."""
    out_dir = str(Path(output_folder).resolve()) if output_folder and output_folder.strip() else None
    if out_dir:
        Path(out_dir).mkdir(parents=True, exist_ok=True)

    total = len(items)
    success_docs = 0
    failed_docs = 0
    errors: list[str] = []
    details: list[dict] = []
    backup_files: list[str] = []
    used_output_names: set[str] = set()

    def _execute_items(word_app=None):
        nonlocal success_docs, failed_docs
        for index, item in enumerate(items, start=1):
            if is_cancelled and is_cancelled():
                break

            doc_detail: dict = {
                "index": index,
                "filename": item.filename,
                "file_path": item.file_path,
                "success": False,
                "replacements": 0,
                "backup_path": "",
                "destination": "",
                "error": "",
            }

            try:
                if not os.path.isfile(item.file_path):
                    raise FileNotFoundError(f"源文件不存在：{item.file_path}")

                target_path = item.file_path
                if out_dir:
                    candidate_name = item.filename
                    stem, suffix = os.path.splitext(candidate_name)
                    counter = 2
                    while candidate_name.casefold() in used_output_names or Path(out_dir, candidate_name).exists():
                        candidate_name = f"{stem} ({counter}){suffix}"
                        counter += 1
                    used_output_names.add(candidate_name.casefold())

                    dest_path = str(Path(out_dir, candidate_name))
                    shutil.copy2(item.file_path, dest_path)
                    target_path = dest_path
                    doc_detail["destination"] = dest_path
                else:
                    if create_backup:
                        backup_path = create_safe_backup(item.file_path)
                        backup_files.append(backup_path)
                        doc_detail["backup_path"] = backup_path
                    doc_detail["destination"] = item.file_path

                # Construct placeholder replacements dict: {{key}} -> value
                replacements_map: dict[str, str] = {}
                for k, v in item.replacements.items():
                    if k:
                        replacements_map[f"{{{{{k}}}}}"] = str(v) if v is not None else ""

                if use_com:
                    doc = None
                    try:
                        doc = word_app.Documents.Open(os.path.abspath(target_path))
                        rep_count = _com_replace_document(doc, replacements_map)
                        doc.Save()
                        doc_detail["replacements"] = rep_count
                    finally:
                        if doc is not None:
                            doc.Close(False)
                else:
                    rep_count = replace_docx_fields(target_path, replacements_map)
                    doc_detail["replacements"] = rep_count

                doc_detail["success"] = True
                success_docs += 1

            except Exception as exc:
                doc_detail["error"] = str(exc)
                doc_detail["success"] = False
                failed_docs += 1
                errors.append(f"{item.filename}: {exc}")
                # Clean up copied output file on error
                if out_dir and doc_detail.get("destination") and Path(doc_detail["destination"]).exists():
                    try:
                        Path(doc_detail["destination"]).unlink()
                    except OSError:
                        pass

            details.append(doc_detail)
            if progress:
                progress(index, total, doc_detail)

    if use_com:
        with WordAutomationSession() as word_app:
            _execute_items(word_app)
    else:
        _execute_items(None)

    return MultiDocBatchResult(
        total_docs=total,
        success_docs=success_docs,
        failed_docs=failed_docs,
        errors=errors,
        details=details,
        backup_files=backup_files,
    )
