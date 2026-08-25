"""File utilities for safe atomic writes, non-overwriting backups, and path safety."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import uuid
import zipfile

from docx import Document


def create_safe_backup(file_path: str | os.PathLike) -> str:
    """Create a backup copy without overwriting existing backups.

    If file.docx.backup exists, creates file.docx.backup (2), file.docx.backup (3), etc.
    Returns the absolute path of the created backup file.
    """
    file_path = Path(file_path).resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"源文件不存在，无法创建备份：{file_path}")

    base_backup = file_path.with_name(file_path.name + ".backup")
    if not base_backup.exists():
        shutil.copy2(file_path, base_backup)
        return str(base_backup)

    counter = 2
    while True:
        candidate = file_path.with_name(f"{file_path.name}.backup ({counter})")
        if not candidate.exists():
            shutil.copy2(file_path, candidate)
            return str(candidate)
        counter += 1


def verify_docx_integrity(file_path: str | os.PathLike) -> bool:
    """Verify that a saved .docx/.docm file is a valid openable ZIP archive containing word/document.xml."""
    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            namelist = zf.namelist()
            if "word/document.xml" not in namelist and "word/document2.xml" not in namelist:
                return False
            # Test CRC on all items
            bad_file = zf.testzip()
            return bad_file is None
    except Exception:
        return False


def atomic_save_docx(doc: Document, target_path: str | os.PathLike) -> None:
    """Save a python-docx Document safely using a validated temporary file and atomic replace.

    1. Saves document to a temporary file in the same directory (same filesystem).
    2. Validates zip structure and document.xml integrity.
    3. Atomically replaces target_path with os.replace.
    4. Cleans up temp file on failure, preserving original target file.
    """
    target_path = Path(target_path).resolve()
    target_dir = target_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    temp_filename = f".tmp_{uuid.uuid4().hex[:8]}_{target_path.name}"
    temp_path = target_dir / temp_filename

    try:
        doc.save(str(temp_path))
        if not verify_docx_integrity(temp_path):
            raise IOError(f"文档保存完整性校验失败，临时文件损坏：{temp_path}")
        # Atomic replacement on same filesystem
        os.replace(temp_path, target_path)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except OSError:
                pass
