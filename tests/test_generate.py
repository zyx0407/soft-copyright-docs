"""冒烟测试：用一份最小配置跑完四个生成器，检查产出确实成立。

只用标准库的 unittest，不需要额外安装任何东西：

    python -m unittest discover -s tests -t .

也可以直接用 pytest 跑（它能识别 unittest 的用例）。
"""

from __future__ import annotations

import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ruanzhu import manual_doc, screenshot_doc, source_doc, spec_doc  # noqa: E402
from ruanzhu.config import ConfigError, ProjectConfig  # noqa: E402

SAMPLE_CODE = '''"""示例模块。"""


def add(a, b):
    """两数相加。"""
    return a + b


def sub(a, b):
    """两数相减。"""
    return a - b
'''


def all_text(path: Path) -> str:
    return "\n".join(p.text for p in Document(str(path)).paragraphs)


class BaseCase(unittest.TestCase):
    """搭一个最小的可跑项目：两个源文件 + 一份配置。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="ruanzhu-test-"))
        src = self.tmp / "src"
        src.mkdir()
        (src / "main.py").write_text(SAMPLE_CODE, encoding="utf-8")
        (src / "helper.py").write_text(
            SAMPLE_CODE.replace("add", "mul"), encoding="utf-8")

        raw = {
            "name": "测试软件",
            "version": "V1.0",
            "output_dir": "out",
            "summary": "这是一个用于测试的软件。",
            "hardware": ["内存：2 GB"],
            "software": ["操作系统：Windows 10"],
            "features": [{"title": "功能甲", "detail": "说明甲。"}],
            "glossary": [{"term": "T", "full": "Term", "desc": "术语说明"}],
            "modules": [
                {"id": "1", "name": "模块甲", "desc": "模块说明。",
                 "steps": ["第一步。", "第二步。"]},
            ],
            "source": {"dirs": ["src"], "include": ["**/*.py"]},
        }
        path = self.tmp / "project.json"
        path.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
        self.project = ProjectConfig.load(path)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


class TestConfig(BaseCase):
    def test_rejects_unknown_field(self):
        path = self.tmp / "bad.json"
        path.write_text('{"name": "x", "nam": "y"}', encoding="utf-8")
        with self.assertRaises(ConfigError):
            ProjectConfig.load(path)

    def test_rejects_bad_json(self):
        path = self.tmp / "bad.json"
        path.write_text("{not json", encoding="utf-8")
        with self.assertRaises(ConfigError):
            ProjectConfig.load(path)

    def test_defaults_are_the_filing_requirements(self):
        self.assertEqual(self.project.source.lines_per_page, 55)
        self.assertEqual(self.project.source.first_pages, 30)
        self.assertEqual(self.project.source.last_pages, 30)

    def test_source_files_are_collected(self):
        names = {p.name for p in self.project.source_files()}
        self.assertEqual(names, {"main.py", "helper.py"})

    def test_output_path_follows_config_location(self):
        self.assertEqual(self.project.output_path(), self.tmp / "out")


class TestPagination(unittest.TestCase):
    def test_paginate(self):
        lines = [f"line {i}" for i in range(100)]
        pages = source_doc.paginate(lines, 55)
        self.assertEqual(len(pages), 2)
        self.assertEqual(len(pages[0]), 55)
        self.assertEqual(len(pages[1]), 45)

    def test_takes_front_and_back(self):
        pages = [["p"] for _ in range(100)]
        selected = source_doc.select_pages(pages, 30, 30)
        self.assertEqual(len(selected), 60)
        self.assertEqual(selected[0], ("front", 1, ["p"]))
        self.assertEqual(selected[29], ("front", 30, ["p"]))
        self.assertEqual(selected[30], ("back", 71, ["p"]))
        self.assertEqual(selected[-1], ("back", 100, ["p"]))

    def test_fewer_pages_than_requested(self):
        pages = [["p"] for _ in range(10)]
        selected = source_doc.select_pages(pages, 30, 30)
        self.assertEqual(len(selected), 10)
        self.assertTrue(all(part == "front" for part, _, _ in selected))


class TestGeneratedDocuments(BaseCase):
    def test_all_four_documents_are_produced(self):
        produced = [
            source_doc.generate(self.project),
            spec_doc.generate(self.project),
            manual_doc.generate(self.project),
            screenshot_doc.generate(self.project),
        ]
        for path in produced:
            self.assertTrue(path.exists(), path)
            self.assertGreater(path.stat().st_size, 5000,
                               f"{path.name} 太小，可能没写进内容")

    def test_source_doc_contents(self):
        text = all_text(source_doc.generate(self.project))
        self.assertIn("测试软件", text)
        self.assertIn("源程序文档", text)
        self.assertIn("# 文件：src/main.py", text)
        self.assertIn("（前部）", text)

    def test_spec_doc_chapters(self):
        text = all_text(spec_doc.generate(self.project))
        self.assertIn("第一章  软件概述", text)
        self.assertIn("第三章  功能模块说明", text)

    def test_manual_doc_steps(self):
        text = all_text(manual_doc.generate(self.project))
        self.assertIn("第三章  功能操作说明", text)
        self.assertIn("步骤 1：", text)

    def test_screenshot_doc_has_index_and_placeholders(self):
        doc = Document(str(screenshot_doc.generate(self.project)))
        self.assertGreaterEqual(len(doc.tables), 2)
        header = [c.text for c in doc.tables[0].rows[0].cells]
        self.assertEqual(header, ["序号", "功能模块", "说明"])

    def test_inline_bold_is_parsed_not_printed(self):
        doc = Document(str(spec_doc.generate(self.project)))
        plain = "\n".join(p.text for p in doc.paragraphs)
        self.assertNotIn("**", plain,
                         "行内粗体标记没有被解析，星号被原样打进了文档")

        feature_paras = [p for p in doc.paragraphs if "功能甲" in p.text]
        self.assertTrue(feature_paras, "没有找到核心功能那一段")
        bold_runs = [r.text for r in feature_paras[0].runs if r.font.bold]
        self.assertIn("功能甲", bold_runs)


if __name__ == "__main__":
    unittest.main(verbosity=2)
