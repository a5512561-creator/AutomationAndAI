# -*- coding: utf-8 -*-
"""parser 單元測試。"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

# 專案根目錄加入 path 以便 import parser
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from parser import Record, parse_line, parse_file


class TestParseLine(unittest.TestCase):
    """parse_line 單行解析測試。"""

    def test_valid_digital_svn(self):
        line = "RL6665 Digital SVN RT\tR8943\t胡定安\t3519497\t通訊網路事業群"
        rec = parse_line(line, 1)
        self.assertIsNotNone(rec)
        self.assertEqual(rec.project, "RL6665")
        self.assertEqual(rec.group, "Digital")
        self.assertTrue(rec.add_svn)
        self.assertEqual(rec.employee_id, "R8943")
        self.assertEqual(rec.line_number, 1)

    def test_valid_analog_svn(self):
        line = "RL6665 Analog SVN RT\tR8067\t王泓閔\t3510572"
        rec = parse_line(line, 2)
        self.assertIsNotNone(rec)
        self.assertEqual(rec.project, "RL6665")
        self.assertEqual(rec.group, "Analog")
        self.assertTrue(rec.add_svn)
        self.assertEqual(rec.employee_id, "R8067")

    def test_no_svn_token(self):
        line = "RL6665 Digital RT\tR1234\t某人"
        rec = parse_line(line, 3)
        self.assertIsNotNone(rec)
        self.assertFalse(rec.add_svn)
        self.assertEqual(rec.employee_id, "R1234")

    def test_empty_line_returns_none(self):
        self.assertIsNone(parse_line("", 4))
        self.assertIsNone(parse_line("   \t  ", 5))
        self.assertIsNone(parse_line("\n", 6))

    def test_no_tab_returns_none(self):
        line = "RL6665 Digital SVN RT R8943"
        rec = parse_line(line, 7)
        self.assertIsNone(rec)

    def test_first_column_only_returns_none(self):
        line = "RL6665 Digital\t"
        rec = parse_line(line, 8)
        self.assertIsNone(rec)

    def test_single_token_first_column_returns_none(self):
        line = "RL6665\tR8943"
        rec = parse_line(line, 9)
        self.assertIsNone(rec)

    def test_employee_id_from_second_column_first_field(self):
        line = "RL6665 Digital SVN RT\tR5539\t梁禮涵\t12345"
        rec = parse_line(line, 10)
        self.assertIsNotNone(rec)
        self.assertEqual(rec.employee_id, "R5539")


class TestParseFile(unittest.TestCase):
    """parse_file 整檔解析測試。"""

    def test_parse_file_two_lines(self):
        content = (
            "RL6665 Digital SVN RT\tR8943\t胡定安\t3519497\n"
            "RL6665 Analog SVN RT\tR8067\t王泓閔\t3510572\n"
        )
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        ) as f:
            f.write(content)
            path = f.name
        try:
            records = parse_file(path)
            self.assertEqual(len(records), 2)
            self.assertEqual(records[0].employee_id, "R8943")
            self.assertEqual(records[0].group, "Digital")
            self.assertEqual(records[1].employee_id, "R8067")
            self.assertEqual(records[1].group, "Analog")
        finally:
            Path(path).unlink(missing_ok=True)

    def test_parse_file_skips_empty_and_invalid(self):
        content = (
            "RL6665 Digital SVN RT\tR8943\t胡定安\n"
            "\n"
            "RL6665 Digital SVN RT R8943\n"  # 無 Tab，應跳過
            "RL6665 Analog SVN RT\tR8067\t王泓閔\n"
        )
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        ) as f:
            f.write(content)
            path = f.name
        try:
            records = parse_file(path)
            self.assertEqual(len(records), 2)
            self.assertEqual(records[0].employee_id, "R8943")
            self.assertEqual(records[1].employee_id, "R8067")
        finally:
            Path(path).unlink(missing_ok=True)

    def test_parse_file_nonexistent_returns_empty(self):
        records = parse_file(Path("/nonexistent/file.txt"))
        self.assertEqual(records, [])


if __name__ == "__main__":
    unittest.main()
