"""Unit tests for the headless replacer_core module."""

import os
from pathlib import Path
import shutil
import sys
import tempfile
import time
import unittest

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from replacer_core import (  # noqa: E402
    compile_search_pattern,
    count_occurrences,
    find_match_contexts,
    get_document_text,
    perform_standard_preview,
    perform_standard_replace,
    replace_in_paragraph_advanced,
    replace_in_table,
    scan_hyperlinks,
    strip_invisible_chars,
)


class ReplacerCoreTests(unittest.TestCase):
    def test_invisible_chars_cleaning_helper(self):
        dirty = "甲\u00ad方\u200b名\u200c称\u200d\u2060\ufeff"
        cleaned = strip_invisible_chars(dirty)
        self.assertEqual(cleaned, "甲方名称")

    def test_invisible_chars_preserved_if_no_match(self):
        doc = Document()
        p = doc.add_paragraph("甲\u00ad方\u200b名\u200c称")
        orig_text = p.runs[0].text

        count = replace_in_paragraph_advanced(p, "乙方", "丙方")
        self.assertEqual(count, 0)
        # Verify run text was NOT mutated or stripped when there's no match
        self.assertEqual(p.runs[0].text, orig_text)

    def test_invisible_chars_matching_and_replacement(self):
        doc = Document()
        p = doc.add_paragraph("甲\u00ad方\u200b公司")
        count = replace_in_paragraph_advanced(p, "甲方", "乙方")
        self.assertEqual(count, 1)
        self.assertEqual(strip_invisible_chars(p.text), "乙方公司")
        self.assertTrue(p.text.startswith("乙方"))

    def test_count_occurrences_modes(self):
        text = "Hello world, hello WORLD. Hello123 World."
        self.assertEqual(count_occurrences(text, "hello", case_sensitive=False), 3)
        self.assertEqual(count_occurrences(text, "Hello", case_sensitive=True), 2)
        self.assertEqual(count_occurrences(text, "Hello", case_sensitive=False, whole_word=True), 2)
        self.assertEqual(count_occurrences(text, "Hello\\d+", use_regex=True), 1)

    def test_multiple_matches_across_independent_runs(self):
        # Paragraph with 3 runs: "x", " / ", "x"
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("x")
        r2 = p.add_run(" / ")
        r3 = p.add_run("x")

        count = replace_in_paragraph_advanced(p, "x", "y")
        self.assertEqual(count, 2)
        self.assertEqual(p.text, "y / y")
        self.assertEqual(r1.text, "y")
        self.assertEqual(r2.text, " / ")
        self.assertEqual(r3.text, "y")

    def test_cross_run_replacement_preserves_unaffected_run_formatting(self):
        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run("Prefix ")
        r2 = p.add_run("BOLD ")
        r2.bold = True
        r3 = p.add_run("ITALIC ")
        r3.italic = True
        r4 = p.add_run("Suffix")
        r4.font.name = "Arial"

        self.assertEqual(p.text, "Prefix BOLD ITALIC Suffix")
        count = replace_in_paragraph_advanced(p, "BOLD ITALIC", "NEW_MIDDLE")
        self.assertEqual(count, 1)
        self.assertEqual(p.text, "Prefix NEW_MIDDLE Suffix")
        self.assertEqual(r1.text, "Prefix ")
        self.assertEqual(r4.text, "Suffix")
        self.assertEqual(r4.font.name, "Arial")

    def test_regex_backreference_expansion(self):
        doc = Document()
        p = doc.add_paragraph("Contact: user123@domain.com")
        count = replace_in_paragraph_advanced(
            p, r"(\w+)@(\w+)\.com", r"\1 at \2 dot com", use_regex=True
        )
        self.assertEqual(count, 1)
        self.assertEqual(p.text, "Contact: user123 at domain dot com")

    def test_replace_in_table_and_nested_table(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "原告：张三"
        table.cell(0, 1).text = "被告：李四"
        table.cell(1, 0).text = "案由：借款"
        table.cell(1, 1).text = "原告住所地：北京"

        count = replace_in_table(table, "原告", "申请人")
        self.assertEqual(count, 2)
        self.assertIn("申请人：张三", table.cell(0, 0).text)
        self.assertIn("申请人住所地：北京", table.cell(1, 1).text)

    def test_zero_matches_does_not_save_or_create_backup(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        with tempfile.TemporaryDirectory() as temp_dir:
            test_doc = Path(temp_dir) / "test_contract.docx"
            shutil.copy2(sample_template, test_doc)
            orig_mtime = test_doc.stat().st_mtime

            time.sleep(0.05)
            res = perform_standard_replace(
                file_paths=[str(test_doc)],
                search_text="不存在的超长字符串123456",
                replace_text="替换文本",
                create_backup=True,
            )
            self.assertEqual(res.total_count, 0)
            self.assertEqual(len(res.backup_files), 0)
            # Verify no backup file created on disk
            self.assertFalse((test_doc.with_name(test_doc.name + ".backup")).exists())
            # Verify file was not written to
            self.assertEqual(test_doc.stat().st_mtime, orig_mtime)

    def test_standard_replace_and_backup(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        with tempfile.TemporaryDirectory() as temp_dir:
            test_doc = Path(temp_dir) / "test_contract.docx"
            shutil.copy2(sample_template, test_doc)

            res = perform_standard_replace(
                file_paths=[str(test_doc)],
                search_text="合同",
                replace_text="协议",
                create_backup=True,
            )
            self.assertGreater(res.total_count, 0)
            self.assertEqual(len(res.backup_files), 1)
            self.assertTrue(Path(res.backup_files[0]).exists())

            # Verify replaced document
            new_text = get_document_text(Document(test_doc))
            self.assertIn("协议", new_text)

    def test_standard_preview(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        res = perform_standard_preview(
            file_paths=[str(sample_template)],
            search_text="合同",
        )
        self.assertGreater(res.total_count, 0)
        self.assertEqual(res.files_with_matches, 1)

    def test_scan_hyperlinks(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        links = scan_hyperlinks([str(sample_template)])
        self.assertEqual(len(links), 1)
        self.assertIsNone(links[0]["error"])
