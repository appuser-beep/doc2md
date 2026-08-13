# -*- coding: utf-8 -*-
"""ZIP 增强：魔数识别、真 ZIP 源码包、伪 ZIP(RAR) 阻断。"""

from __future__ import annotations

import io
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from converter import ConversionError, convert_path, precheck_source  # noqa: E402
from zip_convert import (  # noqa: E402
    convert_zip_to_markdown,
    decode_zip_member_name,
    sniff_archive_kind,
)

SAMPLES = ROOT / "tests" / "zip_cases" / "samples"
OUTPUT = ROOT / "tests" / "zip_cases" / "output"
FAKE_RAR_AS_ZIP = Path(r"D:\Java_SE_HomeWork.zip")


def _write_real_java_zip() -> Path:
    SAMPLES.mkdir(parents=True, exist_ok=True)
    path = SAMPLES / "java_homework.zip"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "src/Hello.java",
            "public class Hello {\n  public static void main(String[] a) {\n"
            '    System.out.println("hi");\n  }\n}\n',
        )
        zf.writestr("README.md", "# Homework\n\n请阅读说明。\n")
        zf.writestr("out/Hello.class", b"\xca\xfe\xba\xbe" + b"\x00" * 40)
        zf.writestr("lib/demo.jar", b"PK\x03\x04fakejar")
    path.write_bytes(buf.getvalue())
    return path


class TestSniff(unittest.TestCase):
    def test_real_zip_magic(self):
        p = _write_real_java_zip()
        self.assertEqual(sniff_archive_kind(p), "zip")

    def test_desktop_fake_zip_is_rar(self):
        if not FAKE_RAR_AS_ZIP.is_file():
            self.skipTest("桌面伪 ZIP 不存在")
        self.assertEqual(sniff_archive_kind(FAKE_RAR_AS_ZIP), "rar")


class TestZipConvert(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        OUTPUT.mkdir(parents=True, exist_ok=True)
        cls.sample = _write_real_java_zip()

    def test_converts_java_and_skips_class(self):
        md = convert_zip_to_markdown(self.sample)
        (OUTPUT / "java_homework.md").write_text(md, encoding="utf-8")
        self.assertIn("Hello.java", md)
        self.assertIn("public class Hello", md)
        self.assertIn("README.md", md)
        self.assertIn("Homework", md)
        self.assertIn("已跳过", md)
        self.assertIn("Hello.class", md)

    def test_fake_rar_raises_clear_error(self):
        if not FAKE_RAR_AS_ZIP.is_file():
            self.skipTest("桌面伪 ZIP 不存在")
        with self.assertRaises(ValueError) as ctx:
            convert_zip_to_markdown(FAKE_RAR_AS_ZIP)
        self.assertIn("RAR", str(ctx.exception))

    def test_convert_path_blocks_fake_rar(self):
        if not FAKE_RAR_AS_ZIP.is_file():
            self.skipTest("桌面伪 ZIP 不存在")
        tip = precheck_source(str(FAKE_RAR_AS_ZIP))
        self.assertIsNotNone(tip)
        self.assertIn("RAR", tip or "")
        with self.assertRaises(ConversionError) as ctx:
            convert_path(str(FAKE_RAR_AS_ZIP), local_only=False)
        self.assertIn("RAR", str(ctx.exception))

    def test_convert_path_real_zip(self):
        md = convert_path(str(self.sample), local_only=False)
        self.assertIn("```java", md)
        self.assertIn("Hello", md)

    def test_decode_gbk_member_name(self):
        # 模拟 zipfile 把 GBK 字节按 cp437 解成的乱码串
        original = "档案整理活动立项申请表.docx"
        mojibake = original.encode("gbk").decode("cp437")
        fixed = decode_zip_member_name(mojibake, flag_bits=0)
        self.assertEqual(fixed, original)
        # UTF-8 标志时保持原样
        self.assertEqual(decode_zip_member_name("ok/file.txt", flag_bits=0x800), "ok/file.txt")

    def test_gbk_named_zip_roundtrip(self):
        path = SAMPLES / "gbk_cn_names.zip"
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            info = zipfile.ZipInfo()
            cn = "档案说明.txt"
            info.filename = cn.encode("gbk").decode("cp437")
            info.flag_bits = 0  # 无 UTF-8 标志
            zf.writestr(info, "档案整理服务活动说明正文\n".encode("utf-8"))
        path.write_bytes(buf.getvalue())
        md = convert_zip_to_markdown(path)
        self.assertIn("档案说明.txt", md)
        self.assertIn("档案整理服务活动", md)
        self.assertNotIn("╡", md)


if __name__ == "__main__":
    unittest.main(verbosity=2)
