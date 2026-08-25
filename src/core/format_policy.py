"""Unified format, mode, and platform capability policies for all operations."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Literal

from platform_adapter.capabilities import CAPABILITIES

OperationType = Literal["replace", "merge", "multi_doc"]
EngineMode = Literal["fast", "full"]

VALID_WORD_EXTENSIONS = {".docx", ".docm", ".doc"}
VALID_TABLE_EXTENSIONS = {".xlsx", ".xlsm", ".csv", ".xltx", ".xltm"}


def validate_file_format_for_mode(
    file_path: str | os.PathLike,
    mode: EngineMode = "fast",
    operation: OperationType = "replace",
) -> tuple[bool, str | None]:
    """Validate whether a given file path is supported under the specified engine mode and OS."""
    path = Path(file_path)
    suffix = path.suffix.lower()

    if not suffix or suffix not in VALID_WORD_EXTENSIONS:
        return False, f"不支持的文件格式：{path.name}。仅支持 .docx, .docm, .doc 格式文档。"

    if suffix == ".doc":
        if CAPABILITIES.is_macos:
            return False, f"macOS 暂不支持旧版 .doc 格式文件：{path.name}，请转换为 .docx 格式。"
        if mode == "fast":
            return False, f"旧版 .doc 格式文件（{path.name}）无法使用快速模式，请在处理选项中选择【完整模式】。"
        if not CAPABILITIES.has_word_com:
            return False, f"处理 .doc 格式文件（{path.name}）需要 Windows 系统并安装 Microsoft Word。"

    return True, None


def validate_batch_file_formats(
    file_paths: list[str],
    mode: EngineMode = "fast",
    operation: OperationType = "replace",
) -> tuple[bool, str | None]:
    """Validate a batch of files against format and capability constraints."""
    if not file_paths:
        return False, "请先添加至少一个 Word 文档。"

    for f in file_paths:
        if not os.path.isfile(f):
            return False, f"文件不存在或无法访问：{os.path.basename(f)}"
        ok, err = validate_file_format_for_mode(f, mode=mode, operation=operation)
        if not ok:
            return False, err

    return True, None
