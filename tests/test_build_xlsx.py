#!/usr/bin/env python3
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "build_xlsx.py")
FIXTURES = os.path.join(ROOT, "tests", "fixtures")


def run_fixture(name, output, *flags):
    return subprocess.run(
        [sys.executable, SCRIPT, os.path.join(FIXTURES, name), output, *flags],
        text=True,
        capture_output=True,
        check=False,
    )


def run_data(data, input_path, output, *flags):
    with open(input_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    return subprocess.run(
        [sys.executable, SCRIPT, input_path, output, *flags],
        text=True,
        capture_output=True,
        check=False,
    )


GOOD_TAIL_10 = [
    "真清爽",
    "四屏压成一屏后，手指不用每天在图标堆里反复绕路好几遍了。",
    "九个常用的刚刚好",
    "微信和相机也算在这九个里面吗？",
    "红点退退",
    "资源库搜索真香！",
    "首页少一屏脑子也跟着安静了一点点",
    "我先试三天…",
    "确实会忘",
    "以前解锁只是回消息，结果总被别的图标拐走半天才回来。",
]


class BuildXlsxTests(unittest.TestCase):
    def test_legacy_golden_fixture_remains_reproducible(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "golden.xlsx")
            result = run_fixture("legacy-golden-40.json", output, "--check")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(os.path.exists(output))

    def test_good_batch_passes_and_xlsx_is_valid(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "good.xlsx")
            result = run_fixture("good-calibration.json", output, "--check")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue(os.path.exists(output))
            with zipfile.ZipFile(output, "r") as workbook:
                self.assertIsNone(workbook.testzip())

    def test_ten_comment_tail_passes_with_target_batch_size_twenty(self):
        with tempfile.TemporaryDirectory() as tmp:
            with open(os.path.join(FIXTURES, "good-calibration.json"), encoding="utf-8") as f:
                data = json.load(f)
            data["comments"].extend(GOOD_TAIL_10)
            output = os.path.join(tmp, "good-tail.xlsx")
            result = run_data(data, os.path.join(tmp, "input.json"), output, "--check")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("最后一批 10 条", result.stdout + result.stderr)
            self.assertTrue(os.path.exists(output))

    def test_tail_shorter_than_ten_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            with open(os.path.join(FIXTURES, "good-calibration.json"), encoding="utf-8") as f:
                data = json.load(f)
            data["comments"].extend(GOOD_TAIL_10[:5])
            output = os.path.join(tmp, "bad-tail.xlsx")
            result = run_data(data, os.path.join(tmp, "input.json"), output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertIn("尾批只有 5 条", result.stdout + result.stderr)
            self.assertFalse(os.path.exists(output))

    def test_failed_check_does_not_create_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "bad.xlsx")
            result = run_fixture("bad-all-punctuated.json", output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertFalse(os.path.exists(output))

    def test_failed_check_does_not_overwrite_existing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "existing.xlsx")
            with open(output, "wb") as f:
                f.write(b"keep-me")
            result = run_fixture("bad-all-punctuated.json", output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            with open(output, "rb") as f:
                self.assertEqual(f.read(), b"keep-me")

    def test_four_identical_endings_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "consecutive.xlsx")
            result = run_fixture("bad-consecutive-ending.json", output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertIn("连续 4 条", result.stdout + result.stderr)
            self.assertFalse(os.path.exists(output))

    def test_generic_batch_without_long_comments_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "generic.xlsx")
            result = run_fixture("bad-generic.json", output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            combined = result.stdout + result.stderr
            self.assertIn("具体长评太少", combined)
            self.assertIn("万能泛评过多", combined)
            self.assertFalse(os.path.exists(output))

    def test_near_duplicate_comments_are_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "duplicate.xlsx")
            result = run_fixture("bad-near-duplicate.json", output, "--check")
            self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
            self.assertIn("高度相似评论", result.stdout + result.stderr)
            self.assertFalse(os.path.exists(output))

    def test_invalid_schema_returns_usage_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = os.path.join(tmp, "invalid.xlsx")
            result = run_fixture("invalid-schema.json", output, "--check")
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertIn("comments 必须是字符串数组", result.stdout + result.stderr)
            self.assertFalse(os.path.exists(output))


if __name__ == "__main__":
    unittest.main()
