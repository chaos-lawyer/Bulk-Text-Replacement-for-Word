"""Pure Python core search & replace engine for Word documents.

Decoupled from GUI frameworks, supporting python-docx fast processing and Windows Word COM automation.
"""

from __future__ import annotations

import os
import re
from typing import Callable, Optional

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.text.paragraph import Paragraph

from core.file_utils import atomic_save_docx, create_safe_backup
from core.models import BatchProcessResult, FileProcessResult
from core.word_com import WordAutomationSession, traverse_story_ranges

# Invisible characters to normalize for clean text comparison
INVISIBLE_CHARS = "\u00ad\u200b\u200c\u200d\u2060\ufeff"
INVISIBLE_CHARS_TABLE = str.maketrans("", "", INVISIBLE_CHARS)


def strip_invisible_chars(text: str) -> str:
    """Remove invisible formatting characters like zero-width spaces and soft hyphens."""
    if not text:
        return ""
    return text.translate(INVISIBLE_CHARS_TABLE)


def compile_search_pattern(
    search_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
) -> re.Pattern:
    """Compile search criteria into a standardized re.Pattern.

    Raises re.error or ValueError on invalid patterns.
    """
    if not search_text:
        raise ValueError("查找内容不能为空。")

    if use_regex:
        flags = 0 if case_sensitive else re.IGNORECASE
        return re.compile(search_text, flags)

    escaped = re.escape(search_text)
    if whole_word:
        escaped = r"\b" + escaped + r"\b"
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(escaped, flags)


def count_occurrences(
    text: str,
    search_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
) -> int:
    """Count how many times search_text occurs in text according to specified rules."""
    if not text or not search_text:
        return 0

    try:
        pattern = compile_search_pattern(search_text, case_sensitive, use_regex, whole_word)
        return len(pattern.findall(text))
    except re.error:
        return 0


def _collect_table_text(table, text_chunks: list[str], seen_cells: set[int] | None = None) -> None:
    """Collect text from table cells, properly deduplicating merged cells."""
    if seen_cells is None:
        seen_cells = set()

    for row in table.rows:
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen_cells:
                continue
            seen_cells.add(tc_id)
            for p in cell.paragraphs:
                if p.text:
                    text_chunks.append(p.text)
            for nested in cell.tables:
                _collect_table_text(nested, text_chunks, seen_cells)


def get_document_text(doc: Document) -> str:
    """Extract all text from paragraphs and tables in a python-docx Document."""
    text_chunks: list[str] = []
    for p in doc.paragraphs:
        if p.text:
            text_chunks.append(p.text)
    seen_cells: set[int] = set()
    for t in doc.tables:
        _collect_table_text(t, text_chunks, seen_cells)
    return "\n".join(text_chunks)


def replace_in_paragraph_advanced(
    paragraph: Paragraph,
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    compiled_pattern: Optional[re.Pattern] = None,
) -> int:
    r"""Replace text in a paragraph using character-to-run position mapping.

    Guarantees:
    1. Preserves individual run formatting (bold, italic, font, color, etc.) across multiple runs.
    2. Supports multiple matches within the same run and across different runs.
    3. Supports regex backreferences (\1, \g<name>).
    4. Does NOT mutate run text or delete invisible characters if no match is found.
    5. Normalizes invisible characters (soft hyphens, zero-width spaces) for robust matching.
    """
    runs = paragraph.runs
    if not runs:
        return 0

    # Build logical text and character-to-run index mapping
    full_text_chars = []
    char_map: list[tuple[int, int]] = []  # index -> (run_index, offset_in_run)

    for run_idx, run in enumerate(runs):
        run_text = run.text or ""
        for offset, ch in enumerate(run_text):
            full_text_chars.append(ch)
            char_map.append((run_idx, offset))

    full_text = "".join(full_text_chars)
    if not full_text:
        return 0

    if compiled_pattern is not None:
        pattern = compiled_pattern
    else:
        try:
            pattern = compile_search_pattern(search_text, case_sensitive, use_regex, whole_word)
        except re.error:
            return 0

    matches_spans: list[tuple[int, int, str]] = []
    matches = list(pattern.finditer(full_text))
    if matches:
        for m in matches:
            if m.start() < m.end():
                if use_regex:
                    try:
                        rep = m.expand(replace_text)
                    except Exception:
                        rep = replace_text
                else:
                    rep = replace_text
                matches_spans.append((m.start(), m.end(), rep))
    elif any(c in INVISIBLE_CHARS for c in full_text):
        # Fallback: match on normalized view without mutating runs if zero matches
        norm_chars = []
        norm_to_raw: list[int] = []
        for raw_idx, ch in enumerate(full_text):
            if ch not in INVISIBLE_CHARS:
                norm_chars.append(ch)
                norm_to_raw.append(raw_idx)
        norm_text = "".join(norm_chars)
        if norm_text:
            for m in pattern.finditer(norm_text):
                if m.start() < m.end():
                    s_raw = norm_to_raw[m.start()]
                    e_raw = norm_to_raw[m.end() - 1] + 1
                    if use_regex:
                        try:
                            rep = m.expand(replace_text)
                        except Exception:
                            rep = replace_text
                    else:
                        rep = replace_text
                    matches_spans.append((s_raw, e_raw, rep))

    if not matches_spans:
        return 0

    # Process non-overlapping matches from back to front (right to left)
    matches_spans.sort(key=lambda x: x[0], reverse=True)
    for start_char, end_char, replacement_str in matches_spans:
        if start_char == end_char:
            continue

        start_run_idx, start_offset = char_map[start_char]
        end_run_idx, end_offset = char_map[end_char - 1]

        if start_run_idx == end_run_idx:
            # Single-run match: edit run in-place
            target_run = runs[start_run_idx]
            orig_text = target_run.text
            target_run.text = orig_text[:start_offset] + replacement_str + orig_text[end_offset + 1:]
        else:
            # Multi-run match:
            # 1. Prefix and replacement inserted into first run
            first_run = runs[start_run_idx]
            first_run.text = first_run.text[:start_offset] + replacement_str

            # 2. Clear intermediate runs entirely
            for mid_idx in range(start_run_idx + 1, end_run_idx):
                runs[mid_idx].text = ""

            # 3. Retain remainder in the last run (preserving last run's formatting)
            last_run = runs[end_run_idx]
            last_run.text = last_run.text[end_offset + 1:]

    return len(matches_spans)


def replace_in_table(
    table,
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    compiled_pattern: Optional[re.Pattern] = None,
    seen_cells: set[int] | None = None,
) -> int:
    """Replace text in a table recursively, deduplicating merged cells."""
    if seen_cells is None:
        seen_cells = set()

    count = 0
    for row in table.rows:
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen_cells:
                continue
            seen_cells.add(tc_id)

            for paragraph in cell.paragraphs:
                count += replace_in_paragraph_advanced(
                    paragraph, search_text, replace_text, case_sensitive, use_regex, whole_word, compiled_pattern
                )
            for nested_table in cell.tables:
                count += replace_in_table(
                    nested_table, search_text, replace_text, case_sensitive, use_regex, whole_word, compiled_pattern, seen_cells
                )
    return count


def find_match_contexts(
    doc: Document,
    search_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    max_contexts: int = 5,
) -> list[str]:
    """Find sample matching sentences/paragraphs with highlighted context."""
    try:
        pattern = compile_search_pattern(search_text, case_sensitive, use_regex, whole_word)
    except re.error:
        return []

    contexts: list[str] = []
    seen_paras: set[str] = set()

    def _extract_from_paragraph(p_text: str):
        if not p_text or len(contexts) >= max_contexts:
            return
        p_clean = strip_invisible_chars(p_text).strip()
        if not p_clean or p_clean in seen_paras:
            return

        m = pattern.search(p_clean)
        if m:
            seen_paras.add(p_clean)
            start_pos = max(0, m.start() - 30)
            end_pos = min(len(p_clean), m.end() + 30)
            prefix = "..." if start_pos > 0 else ""
            suffix = "..." if end_pos < len(p_clean) else ""
            snippet = prefix + p_clean[start_pos:end_pos].replace("\n", " ") + suffix
            contexts.append(snippet)

    for p in doc.paragraphs:
        _extract_from_paragraph(p.text)
        if len(contexts) >= max_contexts:
            break

    if len(contexts) < max_contexts:
        seen_cells: set[int] = set()
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    tc_id = id(cell._tc)
                    if tc_id in seen_cells:
                        continue
                    seen_cells.add(tc_id)
                    for cp in cell.paragraphs:
                        _extract_from_paragraph(cp.text)
                        if len(contexts) >= max_contexts:
                            break

    return contexts


def scan_hyperlinks(file_paths: list[str]) -> list[dict]:
    """Scan Word documents for all embedded hyperlinks using python-docx."""
    results = []
    for path in file_paths:
        filename = os.path.basename(path)
        item = {
            "filename": filename,
            "path": path,
            "count": 0,
            "urls": [],
            "error": None,
        }
        if not os.path.isfile(path):
            item["error"] = "文件不存在"
            results.append(item)
            continue

        try:
            doc = Document(path)
            urls = []
            for rel in doc.part.rels.values():
                if rel.reltype == RT.HYPERLINK:
                    target_url = rel.target_ref
                    if target_url:
                        urls.append(target_url)
            item["count"] = len(urls)
            item["urls"] = urls
        except Exception as exc:
            item["error"] = str(exc)
        results.append(item)
    return results


def perform_standard_preview(
    file_paths: list[str],
    search_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Preview occurrences across files in fast python-docx mode."""
    pattern = compile_search_pattern(search_text, case_sensitive, use_regex, whole_word)
    total_files = len(file_paths)
    batch_result = BatchProcessResult(total_files=total_files)

    for idx, path in enumerate(file_paths):
        if is_cancelled and is_cancelled():
            break

        filename = os.path.basename(path)
        if progress_callback:
            progress_callback(idx + 1, total_files, f"正在预览：{filename}")

        if not os.path.isfile(path):
            batch_result.errors.append(f"{filename}: 文件不存在")
            continue

        try:
            doc = Document(path)
            body_count = 0
            for p in doc.paragraphs:
                if p.text:
                    body_count += len(pattern.findall(p.text))

            table_count = 0
            seen_cells: set[int] = set()
            for t in doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        tc_id = id(cell._tc)
                        if tc_id in seen_cells:
                            continue
                        seen_cells.add(tc_id)
                        for p in cell.paragraphs:
                            if p.text:
                                table_count += len(pattern.findall(p.text))

            file_total = body_count + table_count
            contexts = find_match_contexts(doc, search_text, case_sensitive, use_regex, whole_word)

            detail = FileProcessResult(
                filename=filename,
                path=path,
                body_count=body_count,
                table_count=table_count,
                total=file_total,
                contexts=contexts,
            )
            detail.details.append(f"正文段落匹配：{body_count} 处")
            detail.details.append(f"表格单元格匹配：{table_count} 处")

            batch_result.details.append(detail)
            batch_result.total_count += file_total
            batch_result.files_processed += 1
            if file_total > 0:
                batch_result.files_with_matches += 1

        except Exception as exc:
            batch_result.errors.append(f"{filename}: 预览失败 ({exc})")

    return batch_result


def perform_standard_replace(
    file_paths: list[str],
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    create_backup: bool = True,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Execute text replacement across files in fast python-docx mode with atomic writing and safe backups."""
    pattern = compile_search_pattern(search_text, case_sensitive, use_regex, whole_word)
    total_files = len(file_paths)
    batch_result = BatchProcessResult(total_files=total_files)

    for idx, path in enumerate(file_paths):
        if is_cancelled and is_cancelled():
            break

        filename = os.path.basename(path)
        if progress_callback:
            progress_callback(idx + 1, total_files, f"正在替换：{filename}")

        if not os.path.isfile(path):
            batch_result.errors.append(f"{filename}: 文件不存在")
            continue

        try:
            doc = Document(path)
            body_count = 0
            for p in doc.paragraphs:
                body_count += replace_in_paragraph_advanced(
                    p, search_text, replace_text, case_sensitive, use_regex, whole_word, compiled_pattern=pattern
                )

            table_count = 0
            seen_cells: set[int] = set()
            for t in doc.tables:
                table_count += replace_in_table(
                    t, search_text, replace_text, case_sensitive, use_regex, whole_word, compiled_pattern=pattern, seen_cells=seen_cells
                )

            file_total = body_count + table_count

            # Only create backup and save if there were actual replacements!
            if file_total > 0:
                if create_backup:
                    backup_path = create_safe_backup(path)
                    batch_result.backup_files.append(backup_path)

                atomic_save_docx(doc, path)

            detail = FileProcessResult(
                filename=filename,
                path=path,
                body_count=body_count,
                table_count=table_count,
                total=file_total,
            )
            detail.details.append(f"正文替换：{body_count} 处")
            detail.details.append(f"表格替换：{table_count} 处")

            batch_result.details.append(detail)
            batch_result.total_count += file_total
            batch_result.files_processed += 1
            batch_result.successful_files += 1
            if file_total > 0:
                batch_result.files_with_matches += 1

        except Exception as exc:
            batch_result.errors.append(f"{filename}: 替换失败 ({exc})")

    return batch_result


# ---------------- COM Automation Implementation ----------------


def perform_com_preview(
    file_paths: list[str],
    search_text: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Preview occurrences using Word COM automation in an isolated session."""
    total_files = len(file_paths)
    batch_result = BatchProcessResult(total_files=total_files)

    with WordAutomationSession() as word_app:
        for idx, path in enumerate(file_paths):
            if is_cancelled and is_cancelled():
                break

            filename = os.path.basename(path)
            if progress_callback:
                progress_callback(idx + 1, total_files, f"正在 COM 预览：{filename}")

            if not os.path.isfile(path):
                batch_result.errors.append(f"{filename}: 文件不存在")
                continue

            doc_obj = None
            try:
                abs_path = os.path.abspath(path)
                doc_obj = word_app.Documents.Open(abs_path, ReadOnly=True)

                def _count_story(story_range) -> int:
                    c = 0
                    rng = story_range.Duplicate
                    rng.Find.ClearFormatting()
                    while rng.Find.Execute(
                        FindText=search_text,
                        Replace=0,
                        Forward=True,
                        Wrap=0,
                        MatchCase=case_sensitive,
                        MatchWholeWord=whole_word,
                        MatchWildcards=False,
                    ):
                        c += 1
                    return c

                file_total, warnings = traverse_story_ranges(doc_obj, _count_story)

                detail = FileProcessResult(filename=filename, path=path, total=file_total)
                detail.details.append(f"COM 全文档匹配：{file_total} 处")
                for w in warnings:
                    detail.details.append(f"⚠️ {w}")

                batch_result.details.append(detail)
                batch_result.total_count += file_total
                batch_result.files_processed += 1
                if file_total > 0:
                    batch_result.files_with_matches += 1

            except Exception as exc:
                batch_result.errors.append(f"{filename}: COM 预览失败 ({exc})")
            finally:
                if doc_obj is not None:
                    try:
                        doc_obj.Close(0)
                    except Exception:
                        pass

    return batch_result


def perform_com_replace(
    file_paths: list[str],
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    create_backup: bool = True,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Execute text replacement using Word COM automation in an isolated session."""
    total_files = len(file_paths)
    batch_result = BatchProcessResult(total_files=total_files)

    with WordAutomationSession() as word_app:
        for idx, path in enumerate(file_paths):
            if is_cancelled and is_cancelled():
                break

            filename = os.path.basename(path)
            if progress_callback:
                progress_callback(idx + 1, total_files, f"正在 COM 替换：{filename}")

            if not os.path.isfile(path):
                batch_result.errors.append(f"{filename}: 文件不存在")
                continue

            doc_obj = None
            try:
                abs_path = os.path.abspath(path)

                # Pre-scan count
                doc_obj = word_app.Documents.Open(abs_path, ReadOnly=False)

                def _replace_story(story_range) -> int:
                    c = 0
                    rng = story_range.Duplicate
                    rng.Find.ClearFormatting()
                    rng.Find.Replacement.ClearFormatting()
                    while rng.Find.Execute(
                        FindText=search_text,
                        ReplaceWith=replace_text,
                        Replace=1,  # wdReplaceOne
                        Forward=True,
                        Wrap=0,
                        MatchCase=case_sensitive,
                        MatchWholeWord=whole_word,
                        MatchWildcards=False,
                    ):
                        c += 1
                    return c

                if create_backup:
                    backup_path = create_safe_backup(path)
                    batch_result.backup_files.append(backup_path)

                file_total, warnings = traverse_story_ranges(doc_obj, _replace_story)
                doc_obj.Save()

                detail = FileProcessResult(filename=filename, path=path, total=file_total)
                detail.details.append(f"COM 替换完成：{file_total} 处")
                for w in warnings:
                    detail.details.append(f"⚠️ {w}")

                batch_result.details.append(detail)
                batch_result.total_count += file_total
                batch_result.files_processed += 1
                batch_result.successful_files += 1
                if file_total > 0:
                    batch_result.files_with_matches += 1

            except Exception as exc:
                batch_result.errors.append(f"{filename}: COM 替换失败 ({exc})")
            finally:
                if doc_obj is not None:
                    try:
                        doc_obj.Close(0)
                    except Exception:
                        pass

    return batch_result
