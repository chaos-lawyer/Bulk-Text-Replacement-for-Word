from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

from core.format_policy import validate_batch_file_formats, validate_file_format_for_mode
from platform_adapter.capabilities import CAPABILITIES


class FormatPolicyTests(unittest.TestCase):
    def test_docx_docm_valid_in_fast_mode(self):
        ok, err = validate_file_format_for_mode("contract.docx", mode="fast")
        self.assertTrue(ok)
        self.assertIsNone(err)

        ok, err = validate_file_format_for_mode("template.docm", mode="fast")
        self.assertTrue(ok)
        self.assertIsNone(err)

    def test_doc_rejected_in_fast_mode(self):
        ok, err = validate_file_format_for_mode("legacy.doc", mode="fast")
        self.assertFalse(ok)
        self.assertTrue("无法使用快速模式" in err or "macOS 暂不支持" in err)

    def test_invalid_extension_rejected(self):
        ok, err = validate_file_format_for_mode("data.pdf", mode="fast")
        self.assertFalse(ok)
        self.assertIn("不支持的文件格式", err)
