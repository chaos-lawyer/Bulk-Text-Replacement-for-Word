import hashlib
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from template_merge import (  # noqa: E402
    build_field_mapping,
    build_output_filename,
    extract_template_fields,
    generate_batch,
    load_excel_data,
    normalize_excel_value,
    replace_docx_fields,
    sanitize_filename,
)


SAMPLES = ROOT / "tests" / "samples"


def document_text(path):
    document = Document(path)
    chunks = []
    for part in document.part.package.parts:
        if hasattr(part, "element"):
            for node in part.element.xpath(".//w:t"):
                chunks.append(node.text or "")
    return "".join(chunks)


class TemplateMergeTests(unittest.TestCase):
    def test_scan_excel_and_fields(self):
        fields = extract_template_fields(SAMPLES / "合同模板.docx")
        self.assertEqual(
            fields,
            ["甲方名称", "乙方名称", "合同金额", "签订日期", "管辖法院", "Contract ID"],
        )
        data = load_excel_data(SAMPLES / "合同数据.xlsx")
        self.assertEqual(len(data.rows), 3)
        self.assertEqual(data.rows[0]["签订日期"], "2026-08-20")
        self.assertEqual(data.rows[1]["合同金额"], "2000000")
        self.assertEqual(normalize_excel_value(None), "")

    def test_standard_replace_preserves_split_run_format_and_hyperlink(self):
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "output.docx"
            shutil.copy2(SAMPLES / "合同模板.docx", output)
            count = replace_docx_fields(
                output,
                {
                    "{{甲方名称}}": "A公司",
                    "{{乙方名称}}": "B公司",
                    "{{合同金额}}": "100万元",
                    "{{签订日期}}": "2026-08-20",
                    "{{管辖法院}}": "太原市中级人民法院",
                    "{{Contract ID}}": "CN-001",
                },
            )
            self.assertGreaterEqual(count, 9)
            self.assertNotIn("{{", document_text(output))
            document = Document(output)
            bold = next(p for p in document.paragraphs if p.text.startswith("粗体字段"))
            bold_replacement = next(run for run in bold.runs if "2026-08-20" in run.text)
            self.assertTrue(bold_replacement.bold)
            split = next(p for p in document.paragraphs if p.text.startswith("拆分字段"))
            replaced = next(run for run in split.runs if "太原市" in run.text)
            self.assertEqual(replaced.font.name, "Arial")
            self.assertEqual(replaced.font.size.pt, 13)
            rels = [rel for rel in document.part.rels.values() if rel.reltype.endswith("/hyperlink")]
            self.assertEqual(len(rels), 1)

    def test_batch_generates_three_files_without_changing_template(self):
        template = SAMPLES / "合同模板.docx"
        before = hashlib.sha256(template.read_bytes()).digest()
        data = load_excel_data(SAMPLES / "合同数据.xlsx")
        fields = extract_template_fields(template)
        mapping = build_field_mapping(fields, data.headers)
        with tempfile.TemporaryDirectory() as folder:
            results = generate_batch(
                template, data, mapping, folder,
                "{{甲方名称}}-{{乙方名称}}-合同.docx",
            )
            self.assertTrue(all(result.success for result in results))
            self.assertEqual(
                [result.filename for result in results],
                ["A公司-B公司-合同.docx", "C公司-D公司-合同.docx", "E公司-F公司-合同.docx"],
            )
            for result in results:
                self.assertNotIn("{{", document_text(Path(folder) / result.filename))
        self.assertEqual(before, hashlib.sha256(template.read_bytes()).digest())

    def test_filename_safety_mapping_and_deduplication(self):
        self.assertEqual(sanitize_filename('A/B:合同?.docx'), "A_B_合同_.docx")
        with tempfile.TemporaryDirectory() as folder:
            reserved = set()
            first = build_output_filename("合同.docx", {}, folder, reserved=reserved)
            second = build_output_filename("合同.docx", {}, folder, reserved=reserved)
            self.assertEqual(first, "合同.docx")
            self.assertEqual(second, "合同 (2).docx")


if __name__ == "__main__":
    unittest.main()
