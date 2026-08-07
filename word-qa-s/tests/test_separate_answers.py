#!/usr/bin/env python3
"""Regression tests for question-section detection and answer grouping."""

from __future__ import annotations

import importlib.util
import re
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPT_PATH = SKILL_DIR / "scripts" / "separate_answers.py"
SPEC = importlib.util.spec_from_file_location("separate_answers", SCRIPT_PATH)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}
VAL = f"{{{W_NS}}}val"


def qn(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def p(style: str | None, text: str, rich_xml: str = "") -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return (
        "<w:p>"
        f"{style_xml}"
        f"<w:r><w:t>{text}</w:t></w:r>{rich_xml}"
        "</w:p>"
    )


def toc_xml() -> str:
    return f'''<w:sdt>
  <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
  <w:sdtContent>
    <w:p><w:pPr><w:pStyle w:val="TOC"/></w:pPr><w:r><w:t>目录</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>旧目录项</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  </w:sdtContent>
</w:sdt>'''


def make_document(path: Path, paragraphs: list[tuple], *, include_toc: bool = False) -> None:
    body = (toc_xml() if include_toc else "") + "".join(p(*paragraph) for paragraph in paragraphs)
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:m="{M_NS}"><w:body>{body}<w:sectPr/></w:body></w:document>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:style w:type="paragraph" w:styleId="h1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="h2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="h3"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="TOC"><w:name w:val="TOC Heading"/></w:style>
  <w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/></w:style>
</w:styles>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", rels)
        archive.writestr("word/_rels/document.xml.rels", document_rels)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)


def text_and_styles(path: Path) -> list[tuple[str | None, str]]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    result = []
    for paragraph in root.findall(".//w:body/w:p", NS):
        pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
        style = pstyle.get(f"{{{W_NS}}}val") if pstyle is not None else None
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))
        result.append((style, text))
    return result


def paragraphs_with_objects(path: Path) -> list[tuple[str, bool]]:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return [
        (
            "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)),
            paragraph.find(".//w:object", NS) is not None,
        )
        for paragraph in root.findall(".//w:body/w:p", NS)
    ]


def object_count(path: Path) -> int:
    root = package_xml(path, "word/document.xml")
    return len(root.findall(".//w:object", NS))


def package_xml(path: Path, name: str) -> ET.Element:
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read(name))


def toc_nodes(path: Path) -> list[ET.Element]:
    root = package_xml(path, "word/document.xml")
    return [
        node
        for node in root.findall(".//w:sdt", NS)
        if node.find("./w:sdtPr/w:docPartObj/w:docPartGallery[@w:val='Table of Contents']", NS)
        is not None
    ]


class SeparateAnswersTests(unittest.TestCase):
    def separate(self, paragraphs: list[tuple[str | None, str]]) -> list[tuple[str | None, str]]:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, paragraphs)
        MODULE.separate_answers(source, output, force=False)
        return text_and_styles(output)

    def test_numbered_chapter_example_creates_one_chapter_answer_area(self) -> None:
        result = self.separate([
            ("h1", "第一章 概述"),
            ("h2", "4. 本章例题"),
            ("h3", "4.1 选择题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案一"),
            ("h3", "4.2 简答题"),
            (None, "2．第二题"),
            (None, "【万人教育解析】答案二"),
            ("h1", "第二章 后续内容"),
        ])
        texts = [text for _, text in result]
        self.assertEqual(texts.count("5. 本章例题答案"), 1)
        self.assertEqual(result[texts.index("5. 本章例题答案")][0], "h2")
        self.assertLess(texts.index("5. 本章例题答案"), texts.index("第二章 后续内容"))
        self.assertEqual(texts.count("4.1 选择题"), 1)
        self.assertEqual(texts.count("5.1 选择题"), 1)
        self.assertEqual(texts.count("4.2 简答题"), 1)
        self.assertEqual(texts.count("5.2 简答题"), 1)
        self.assertNotIn("【万人教育解析】答案一", texts)
        self.assertIn("1.【万人教育解析】答案一", texts)
        self.assertIn("2.【万人教育解析】答案二", texts)

    def test_question_type_children_of_a_section_create_one_section_answer_area(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "1.1 基础概念"),
            ("h3", "选择题"),
            (None, "1、第一题"),
            (None, "【万人教育解析】答案一"),
            ("h3", "简答题"),
            (None, "2. 第二题"),
            (None, "【万人教育解析】答案二"),
            ("h2", "1.2 下一节"),
        ])
        texts = [text for _, text in result]
        self.assertEqual(texts.count("2.1 基础概念（题目答案）"), 1)
        self.assertEqual(result[texts.index("2.1 基础概念（题目答案）")][0], "h2")
        self.assertLess(texts.index("2.1 基础概念（题目答案）"), texts.index("1.2 下一节"))
        self.assertEqual(texts.count("选择题"), 2)
        self.assertEqual(texts.count("简答题"), 2)

    def test_chapter_without_question_keyword_uses_marker_evidence_to_create_answer_area(self) -> None:
        result = self.separate([
            ("h1", "第二章 8086 CPU的结构与功能"),
            ("h2", "2.1 8086/8088微处理器"),
            ("h3", "8086CPU的内部结构"),
            (None, "1、简要解释CPU是什么？"),
            (None, "【万人教育解析】"),
            (None, "它是计算机的核心部件。"),
            (None, "2、8086/8088CPU内部有哪些寄存器？"),
            (None, "【万人教育解析】有14个16位寄存器。"),
            ("h1", "第三章 8086CPU 指令系统"),
        ])
        texts = [text for _, text in result]
        answer_title = "第二章 8086 CPU的结构与功能（题目答案）"
        answer_index = texts.index(answer_title)
        self.assertLess(answer_index, texts.index("第三章 8086CPU 指令系统"))
        self.assertNotIn("【万人教育解析】", texts[:answer_index])
        self.assertIn("1.【万人教育解析】", texts[answer_index:])
        self.assertIn("它是计算机的核心部件。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】有14个16位寄存器。", texts[answer_index:])
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 6],
            [
                "2.1 8086/8088微处理器",
                "8086CPU的内部结构",
                "1.【万人教育解析】",
                "它是计算机的核心部件。",
                "2.【万人教育解析】有14个16位寄存器。",
            ],
        )

    def test_inferred_chapter_scope_requires_consistent_marker_bounded_pairs(self) -> None:
        result = self.separate([
            ("h1", "第五章 混排内容"),
            (None, "1、第一题"),
            (None, "【万人教育解析】答案一"),
            (None, "2、第二题"),
            (None, "没有答案标记的正文"),
            (None, "3、第三题"),
            (None, "【万人教育解析】答案三"),
        ])
        texts = [text for _, text in result]
        self.assertNotIn("第五章 混排内容（题目答案）", texts)
        self.assertIn("【万人教育解析】答案一", texts)
        self.assertIn("【万人教育解析】答案三", texts)
        self.assertIn("需人工复核", texts)

    def test_consistent_section_is_separated_when_its_chapter_is_mixed(self) -> None:
        result = self.separate([
            ("h1", "第十二章 模数和数模转换"),
            ("h2", "12.1 基础概念"),
            (None, "1、第一题"),
            (None, "答：答案一"),
            (None, "2、第二题"),
            (None, "答：答案二"),
            ("h2", "程序应用"),
            (None, "3、第三题"),
            (None, "没有答案标记的正文"),
        ])
        texts = [text for _, text in result]
        answer_title = "12.1 基础概念（题目答案）"
        answer_index = texts.index(answer_title)
        self.assertLess(answer_index, texts.index("程序应用"))
        self.assertIn("1.【万人教育解析】答案一", texts[answer_index:])
        self.assertIn("2.【万人教育解析】答案二", texts[answer_index:])
        self.assertNotIn("答：答案一", texts[:answer_index])

    def test_answer_marker_run_uses_hwxw_font(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, [
            ("h1", "第一章 题目"),
            (None, "1、第一题"),
            (None, "答：答案一"),
        ])
        MODULE.separate_answers(source, output, force=False)
        root = package_xml(output, "word/document.xml")
        answer_paragraph = next(
            paragraph
            for paragraph in root.findall(".//w:body/w:p", NS)
            if "1.【万人教育解析】答案一" == "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))
        )
        marker_run = next(
            run for run in answer_paragraph.findall("w:r", NS)
            if "【万人教育解析】" == "".join(node.text or "" for node in run.findall("w:t", NS))
        )
        fonts = marker_run.find("w:rPr/w:rFonts", NS)
        self.assertIsNotNone(fonts)
        self.assertEqual("华文新魏", fonts.get(qn("eastAsia")))

    def test_question_number_does_not_use_answer_marker_font(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, [
            ("h1", "第一章 题目"),
            (None, "1、第一题"),
            (None, "答：答案一"),
        ])
        MODULE.separate_answers(source, output, force=False)
        root = package_xml(output, "word/document.xml")
        answer_paragraph = next(
            paragraph
            for paragraph in root.findall(".//w:body/w:p", NS)
            if "1.【万人教育解析】答案一" == "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))
        )
        number_run = next(
            run for run in answer_paragraph.findall("w:r", NS)
            if "1." == "".join(node.text or "" for node in run.findall("w:t", NS))
        )
        fonts = number_run.find("w:rPr/w:rFonts", NS)
        self.assertTrue(fonts is None or fonts.get(qn("eastAsia")) != "华文新魏")

    def test_markerless_explanatory_answer_is_moved_when_no_options_are_present(self) -> None:
        result = self.separate([
            ("h1", "第十二章 模数和数模转换"),
            (None, "1、第一题"),
            (None, "这是第一题的完整解释。"),
            (None, "补充说明。"),
            (None, "2、第二题"),
            (None, "答：答案二"),
            (None, "3、第三题"),
            (None, "解：答案三"),
            (None, "4、第四题"),
            (None, "解析：答案四"),
            (None, "5、第五题"),
            (None, "【万人教育解析】答案五"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第十二章 模数和数模转换（题目答案）")
        self.assertNotIn("这是第一题的完整解释。", texts[:answer_index])
        self.assertIn("1.【万人教育解析】这是第一题的完整解释。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】答案二", texts[answer_index:])

    def test_answer_and_next_question_joined_in_one_paragraph_are_split_before_separation(self) -> None:
        result = self.separate([
            ("h1", "第五章 总线"),
            (None, "1、第一题"),
            (None, "【万人教育解析】B。2、第二题"),
            (None, "【万人教育解析】A。"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第五章 总线（题目答案）")
        self.assertIn("2、第二题", texts[:answer_index])
        self.assertNotIn("【万人教育解析】B。2、第二题", texts[:answer_index])
        self.assertIn("1.【万人教育解析】B。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】A。", texts[answer_index:])

    def test_alias_answer_and_next_question_joined_in_one_paragraph_are_split(self) -> None:
        result = self.separate([
            ("h1", "第五章 总线"),
            (None, "1、第一题"),
            (None, "答：A。2、第二题"),
            (None, "解答：B。"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第五章 总线（题目答案）")
        self.assertIn("2、第二题", texts[:answer_index])
        self.assertNotIn("答：A。2、第二题", texts[:answer_index])
        self.assertIn("1.【万人教育解析】A。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】B。", texts[answer_index:])
        self.assertNotIn("需人工复核", texts)

    def test_ascii_period_after_answer_can_start_the_next_question(self) -> None:
        result = self.separate([
            ("h1", "第五章 总线"),
            (None, "1、第一题"),
            (None, "答：A.2、第二题"),
            (None, "解答：B。"),
        ])
        texts = [text for _, text in result]
        self.assertIn("2、第二题", texts)
        self.assertIn("1.【万人教育解析】A.", texts)
        self.assertIn("2.【万人教育解析】B。", texts)

    def test_type_heading_joined_to_first_question_is_split(self) -> None:
        result = self.separate([
            ("h1", "第五章 题目"),
            (None, "三、填空题1. 第一题______"),
            (None, "答：答案一"),
            (None, "2. 第二题______"),
            (None, "解：答案二"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第五章 题目答案")
        self.assertIn("三、填空题", texts[:answer_index])
        self.assertIn("1. 第一题______", texts[:answer_index])
        self.assertIn("1.【万人教育解析】答案一", texts[answer_index:])

    def test_split_joined_answer_preserves_one_embedded_object(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        embedded_object = "<w:r><w:object><w:objectEmbed/></w:object></w:r>"
        make_document(source, [
            ("h1", "第五章 总线"),
            (None, "1、第一题"),
            (None, "【万人教育解析】A。2、第二题", embedded_object),
            (None, "【万人教育解析】B。"),
        ])
        self.assertEqual(1, object_count(source))
        MODULE.separate_answers(source, output, force=False)
        self.assertEqual(1, object_count(output))
        texts = [text for _, text in text_and_styles(output)]
        self.assertIn("1.【万人教育解析】A。", texts)
        self.assertIn("2.【万人教育解析】B。", texts)

    def test_short_chinese_markerless_answer_is_moved(self) -> None:
        result = self.separate([
            ("h1", "第十三章 DMA"),
            (None, "1、第一题"),
            (None, "这是答案。"),
            (None, "补充说明。"),
            (None, "2、第二题"),
            (None, "答：答案二"),
            (None, "3、第三题"),
            (None, "解：答案三"),
            (None, "4、第四题"),
            (None, "解析：答案四"),
            (None, "5、第五题"),
            (None, "【万人教育解析】答案五"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第十三章 DMA（题目答案）")
        self.assertIn("1.【万人教育解析】这是答案。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】答案二", texts[answer_index:])

    def test_range_answer_is_expanded_under_copied_section_and_type_headings(self) -> None:
        result = self.separate([
            ("h1", "第一章 微型计算机基础题目"),
            (None, "1.2、1.3、1.4微机中的数制与码制"),
            (None, "一、选择题"),
            (None, "1、第一题"),
            (None, "2、第二题"),
            (None, "3、第三题"),
            (None, "4、第四题"),
            (None, "5、第五题"),
            (None, "【万人教育解析】"),
            (None, "1-5：BAAAD"),
            ("h1", "第二章 后续内容"),
        ])
        texts = [text for _, text in result]
        answer_title = "第一章 微型计算机基础题目答案"
        answer_index = texts.index(answer_title)
        self.assertLess(answer_index, texts.index("第二章 后续内容"))
        self.assertEqual(texts.count("1.2、1.3、1.4微机中的数制与码制"), 2)
        self.assertEqual(texts.count("一、选择题"), 2)
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 8],
            [
                "1.2、1.3、1.4微机中的数制与码制",
                "一、选择题",
                "1.【万人教育解析】B",
                "2.【万人教育解析】A",
                "3.【万人教育解析】A",
                "4.【万人教育解析】A",
                "5.【万人教育解析】D",
            ],
        )
        self.assertNotIn("1-5：BAAAD", texts[:answer_index])
        self.assertNotIn("需人工复核", texts)

    def test_alternative_answer_prefixes_are_normalized(self) -> None:
        for prefix in ("答：", "答案:", "解：", "解答:", "解析："):
            with self.subTest(prefix=prefix):
                result = self.separate([
                    ("h1", "第一章 题目"),
                    (None, "1、第一题"),
                    (None, f"{prefix}B"),
                ])
                texts = [text for _, text in result]
                self.assertIn("1.【万人教育解析】B", texts)
                self.assertFalse(any(text.startswith(prefix) for text in texts))
                self.assertNotIn("需人工复核", texts)

    def test_explanation_prefix_followed_by_final_marker_is_moved_as_one_answer_block(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            (None, "1、第一题"),
            (None, "解析：这是解题思路。"),
            (None, "【万人教育解析】A"),
            (None, "2、第二题"),
            (None, "答：答案二"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 题目答案")
        self.assertNotIn("解析：这是解题思路。", texts[:answer_index])
        self.assertNotIn("答：答案二", texts[:answer_index])
        self.assertIn("1.【万人教育解析】这是解题思路。", texts[answer_index:])
        self.assertIn("A", texts[answer_index:])
        self.assertNotIn("【万人教育解析】A", texts[answer_index:])
        self.assertIn("2.【万人教育解析】答案二", texts[answer_index:])

    def test_analysis_alias_and_final_answer_are_moved_as_one_block(self) -> None:
        result = self.separate([
            ("h1", "第十三章 DMA"),
            (None, "1、第一题"),
            (None, "分析：这是解题思路。"),
            (None, "解答：C。"),
            (None, "2、第二题"),
            (None, "题目分析：这是第二题思路。"),
            (None, "答案：B。"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第十三章 DMA（题目答案）")
        self.assertIn("1.【万人教育解析】这是解题思路。", texts[answer_index:])
        self.assertIn("C。", texts[answer_index:])
        self.assertIn("2.【万人教育解析】这是第二题思路。", texts[answer_index:])
        self.assertIn("B。", texts[answer_index:])
        self.assertFalse(any(re.match(r"^(?:分析|题目分析|解答|答案)[：:]", text) for text in texts))

    def test_numeric_data_before_inline_marker_is_not_split_as_a_question(self) -> None:
        result = self.separate([
            ("h1", "第六章 存储器设计"),
            (None, "13.若系统分别使用512K*8、1K*4、16K*8、64K*1的RAM，各需要多少条地址线？【万人教育解析】19、10、14、16。"),
            (None, "14.下一题"),
            (None, "答：答案十四"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第六章 存储器设计（题目答案）")
        self.assertIn("13.若系统分别使用512K*8、1K*4、16K*8、64K*1的RAM，各需要多少条地址线？", texts[:answer_index])
        self.assertIn("13.【万人教育解析】19、10、14、16。", texts[answer_index:])
        self.assertIn("14.【万人教育解析】答案十四", texts[answer_index:])

    def test_inline_question_and_answer_marker_are_split_and_normalized(self) -> None:
        result = self.separate([
            ("h1", "第六章 存储器设计"),
            (None, "1、第一题【万人教育解析】答案一"),
            (None, "2、第二题【万人教育解析】答案二"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第六章 存储器设计（题目答案）")
        self.assertIn("1、第一题", texts[:answer_index])
        self.assertIn("2、第二题", texts[:answer_index])
        self.assertFalse(any("【万人教育解析】" in text for text in texts[:answer_index]))
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 3],
            ["1.【万人教育解析】答案一", "2.【万人教育解析】答案二"],
        )

    def test_question_text_ending_in_please_answer_is_not_an_inline_answer_prefix(self) -> None:
        result = self.separate([
            ("h1", "第六章 存储器设计"),
            (None, "1、请回答："),
            (None, "解：答案一"),
            (None, "2、请回答："),
            (None, "答：答案二"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第六章 存储器设计（题目答案）")
        self.assertIn("1、请回答：", texts[:answer_index])
        self.assertIn("2、请回答：", texts[:answer_index])
        self.assertIn("1.【万人教育解析】答案一", texts[answer_index:])
        self.assertIn("2.【万人教育解析】答案二", texts[answer_index:])
        self.assertNotIn("1.【万人教育解析】", texts[answer_index:])

    def test_question_number_touching_a_numeric_acronym_is_not_treated_as_a_decimal_formula(self) -> None:
        result = self.separate([
            ("h1", "第七章 接口技术"),
            (None, "1、第一题"),
            (None, "答：答案一"),
            (None, "6.8086 CPU 的问题是什么？"),
            (None, "答：答案六"),
            (None, "7. 8086 CPU 的第二个问题是什么？"),
            (None, "答：答案七"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第七章 接口技术（题目答案）")
        self.assertIn("6.8086 CPU 的问题是什么？", texts[:answer_index])
        self.assertIn("7. 8086 CPU 的第二个问题是什么？", texts[:answer_index])
        self.assertIn("6.【万人教育解析】答案六", texts[answer_index:])
        self.assertIn("7.【万人教育解析】答案七", texts[answer_index:])

    def test_decimal_with_voltage_unit_is_not_a_question_number(self) -> None:
        self.assertIsNone(MODULE.QUESTION_RE.match("2.5V/0.019V=128=80H"))

    def test_subquestion_answer_labels_form_one_answer_block(self) -> None:
        result = self.separate([
            ("h1", "第十二章 题目"),
            (None, "1、综合程序题"),
            (None, "题（1）解：第一问答案"),
            (None, "题（2）（3）解：第二、三问答案"),
            (None, "分析：补充说明"),
            (None, "2、下一题"),
            (None, "答：答案二"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第十二章 题目答案")
        self.assertIn("1.【万人教育解析】第一问答案", texts[answer_index:])
        self.assertIn("第二、三问答案", texts[answer_index:])
        self.assertIn("补充说明", texts[answer_index:])
        self.assertNotIn("需人工复核", texts)

    def test_unmatched_answer_alias_is_normalized_in_place_and_marked_for_review(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            (None, "1、题干"),
            (None, "答：答案甲"),
            (None, "答：答案乙"),
        ])
        texts = [text for _, text in result]
        self.assertNotIn("答：答案甲", texts)
        self.assertNotIn("答：答案乙", texts)
        self.assertIn("【万人教育解析】答案甲", texts)
        self.assertIn("【万人教育解析】答案乙", texts)
        self.assertIn("需人工复核", texts)

    def test_duplicate_question_number_does_not_hide_missing_answer(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            (None, "1、已经有答案的题目"),
            (None, "答：A"),
            (None, "1、同一标题下没有答案的另一题"),
            (None, "普通说明文字"),
        ])
        texts = [text for _, text in result]
        self.assertIn("1、同一标题下没有答案的另一题", texts)
        self.assertIn("需人工复核", texts)
        self.assertTrue(any("第 1 题未找到" in text for text in texts))

    def test_unnumbered_first_question_after_heading_resets_to_one(self) -> None:
        result = self.separate([
            ("h1", "第六章 题目"),
            (None, "1、第一题"),
            (None, "答：答案一"),
            ("h2", "程序应用"),
            (None, "试编写一段程序完成转换。"),
            (None, "解：程序答案"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第六章 题目答案")
        self.assertIn("1、试编写一段程序完成转换。", texts[:answer_index])
        self.assertNotIn("2、试编写一段程序完成转换。", texts[:answer_index])
        self.assertEqual(2, texts[answer_index:].count("1.【万人教育解析】答案一") + texts[answer_index:].count("1.【万人教育解析】程序答案"))

    def test_unnumbered_question_in_same_heading_increments_previous_number(self) -> None:
        result = self.separate([
            ("h1", "第六章 题目"),
            ("h2", "程序应用"),
            (None, "1、第一题"),
            (None, "答：答案一"),
            (None, "试编写一段程序完成转换。"),
            (None, "解：程序答案"),
        ])
        texts = [text for _, text in result]
        self.assertIn("2、试编写一段程序完成转换。", texts)
        self.assertIn("2.【万人教育解析】程序答案", texts)

    def test_bracketed_examples_are_numbered_and_separated(self) -> None:
        result = self.separate([
            ("h1", "第一章 状态空间表达式"),
            ("h2", "1.1 基本概念"),
            ("h3", "知识点1：标准型"),
            (None, "【例】第一个例题"),
            (None, "【万人教育解析】第一个答案"),
            (None, "【例】第二个例题"),
            (None, "【万人教育解析】第二个答案"),
            ("h1", "第二章 后续内容"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 状态空间表达式（题目答案）")
        self.assertIn("1、【例】第一个例题", texts[:answer_index])
        self.assertIn("2、【例】第二个例题", texts[:answer_index])
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 5],
            [
                "1.1 基本概念",
                "知识点1：标准型",
                "1.【万人教育解析】第一个答案",
                "2.【万人教育解析】第二个答案",
            ],
        )
        self.assertNotIn("【万人教育解析】第一个答案", texts[:answer_index])
        self.assertNotIn("需人工复核", texts)

    def test_bracketed_example_keeps_numbered_solution_steps_in_its_answer(self) -> None:
        result = self.separate([
            ("h1", "第一章 状态空间表达式"),
            ("h2", "1.1 基本概念"),
            (None, "【例】第一个例题"),
            (None, "【万人教育解析】解题过程"),
            (None, "1. 第一步"),
            (None, "2. 第二步"),
            (None, "【例】第二个例题"),
            (None, "【万人教育解析】第二题答案"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 状态空间表达式（题目答案）")
        self.assertIn("1.【万人教育解析】解题过程", texts[answer_index:])
        self.assertIn("1. 第一步", texts[answer_index:])
        self.assertIn("2. 第二步", texts[answer_index:])
        self.assertIn("2.【万人教育解析】第二题答案", texts[answer_index:])
        self.assertNotIn("需人工复核", texts)

    def test_bracketed_example_keeps_explicit_subanswers_in_one_answer_block(self) -> None:
        result = self.separate([
            ("h1", "第一章 状态空间表达式"),
            ("h2", "1.1 基本概念"),
            (None, "【例】判断下列三个系统"),
            (None, "（1）"),
            (None, "【万人教育解析】第一问结论"),
            (None, "（2）"),
            (None, "【万人教育解析】第二问结论"),
            (None, "【例】下一题"),
            (None, "【万人教育解析】下一题答案"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 状态空间表达式（题目答案）")
        self.assertIn("1.【万人教育解析】第一问结论", texts[answer_index:])
        self.assertIn("（2）", texts[answer_index:])
        self.assertIn("第二问结论", texts[answer_index:])
        self.assertIn("2.【万人教育解析】下一题答案", texts[answer_index:])
        self.assertNotIn("需人工复核", texts)

    def test_bracketed_question_joined_to_preceding_prose_is_split(self) -> None:
        result = self.separate([
            ("h1", "第一章 状态空间表达式"),
            ("h2", "1.1 基本概念"),
            (None, "计算方法如下。【题1】求第一个结果"),
            (None, "【万人教育解析】第一个答案"),
            (None, "【题2】求第二个结果"),
            (None, "【万人教育解析】第二个答案"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 状态空间表达式（题目答案）")
        self.assertIn("计算方法如下。", texts[:answer_index])
        self.assertIn("【题1】求第一个结果", texts[:answer_index])
        self.assertIn("1.【万人教育解析】第一个答案", texts[answer_index:])
        self.assertIn("2.【万人教育解析】第二个答案", texts[answer_index:])
        self.assertNotIn("需人工复核", texts)

    def test_assembly_answer_blank_is_not_restored_as_a_new_question(self) -> None:
        result = self.separate([
            ("h1", "第六章 题目"),
            ("h2", "程序应用"),
            (None, "1、补全下列汇编程序"),
            (None, "答：程序如下"),
            (None, "________"),
            (None, "解析：横线是程序中的待填步骤"),
        ])
        texts = [text for _, text in result]
        self.assertFalse(any(re.match(r"^2\s*[、.]\s*_+", text) for text in texts))

    def test_markerless_prose_does_not_infer_an_ordinary_chapter_scope(self) -> None:
        result = self.separate([
            ("h1", "附录 芯片资料"),
            (None, "1、器件规格"),
            (None, "该器件包含多组工作状态，不同模式之间需要分别查阅。"),
            (None, "各项参数随工作条件发生变化。"),
            (None, "2、时序资料"),
            (None, "不同工作方式下的时序关系并不相同，使用时应查阅手册。"),
            (None, "相关参数列在后续资料中。"),
        ])
        texts = [text for _, text in result]
        self.assertNotIn("附录 芯片资料（题目答案）", texts)
        self.assertIn("该器件包含多组工作状态，不同模式之间需要分别查阅。", texts)
        self.assertIn("不同工作方式下的时序关系并不相同，使用时应查阅手册。", texts)

    def test_unsafe_range_answer_is_left_for_manual_review(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            (None, "1、第一题"),
            (None, "2、第二题"),
            (None, "【万人教育解析】"),
            (None, "1-3：ABC"),
        ])
        texts = [text for _, text in result]
        self.assertIn("【万人教育解析】", texts)
        self.assertIn("1-3：ABC", texts)
        self.assertIn("需人工复核", texts)

    def test_multiple_consecutive_ranges_share_one_answer_marker(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            (None, "1、第一题"),
            (None, "2、第二题"),
            (None, "3、第三题"),
            (None, "4、第四题"),
            (None, "5、第五题"),
            (None, "6、第六题"),
            (None, "【万人教育解析】"),
            (None, "1-3：ABC"),
            (None, "4-6：DEF"),
        ])
        texts = [text for _, text in result]
        self.assertEqual(
            [text for text in texts if text.endswith(("A", "B", "C", "D", "E", "F"))],
            [
                "1.【万人教育解析】A",
                "2.【万人教育解析】B",
                "3.【万人教育解析】C",
                "4.【万人教育解析】D",
                "5.【万人教育解析】E",
                "6.【万人教育解析】F",
            ],
        )
        self.assertNotIn("需人工复核", texts)

    def test_range_answer_uses_the_nearest_heading_chain_when_numbers_restart(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            ("h2", "1.1 旧小节"),
            ("h3", "选择题"),
            (None, "1、旧题"),
            (None, "【万人教育解析】Z"),
            (None, "1.2、1.3、1.4微机中的数制与码制"),
            (None, "一、选择题"),
            (None, "1、第一题"),
            (None, "2、第二题"),
            (None, "3、第三题"),
            (None, "4、第四题"),
            (None, "5、第五题"),
            (None, "【万人教育解析】"),
            (None, "1-5：BAAAD"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 题目答案")
        target_index = texts.index("1.2、1.3、1.4微机中的数制与码制", answer_index + 1)
        self.assertEqual(
            texts[target_index : target_index + 7],
            [
                "1.2、1.3、1.4微机中的数制与码制",
                "一、选择题",
                "1.【万人教育解析】B",
                "2.【万人教育解析】A",
                "3.【万人教育解析】A",
                "4.【万人教育解析】A",
                "5.【万人教育解析】D",
            ],
        )

    def test_same_level_section_heading_is_copied_before_its_type_heading(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            ("h3", "1.2、1.3、1.4微机中的数制与码制"),
            ("h3", "一、选择题"),
            (None, "1、第一题"),
            (None, "2、第二题"),
            (None, "【万人教育解析】"),
            (None, "1-2：BA"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 题目答案")
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 5],
            [
                "1.2、1.3、1.4微机中的数制与码制",
                "一、选择题",
                "1.【万人教育解析】B",
                "2.【万人教育解析】A",
            ],
        )

    def test_numbered_answer_table_is_split_and_removed_from_question_area(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            ("h2", "二、填空题"),
            (None, "1、第一题______"),
            (None, "2、第二题______"),
            (None, "3、第三题______"),
            (None, "【万人教育解析】"),
            (None, "1、答案甲"),
            (None, "2、答案乙"),
            (None, "3、答案丙"),
            ("h2", "三、计算题"),
        ])
        texts = [text for _, text in result]
        answer_index = texts.index("第一章 题目答案")
        self.assertNotIn("1、答案甲", texts[:answer_index])
        self.assertNotIn("2、答案乙", texts[:answer_index])
        self.assertEqual(
            texts[answer_index + 1 : answer_index + 5],
            [
                "二、填空题",
                "1.【万人教育解析】答案甲",
                "2.【万人教育解析】答案乙",
                "3.【万人教育解析】答案丙",
            ],
        )

    def test_math_answer_marker_and_equations_are_moved_as_one_answer_block(self) -> None:
        math_marker = "<m:oMath><m:r><m:t>【万人教育解析】</m:t></m:r></m:oMath>"
        equation = "<m:oMath><m:r><m:t>11010.1=26.5</m:t></m:r></m:oMath>"
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, [
            ("h1", "第一章 题目"),
            (None, "6、计算"),
            (None, "", math_marker),
            (None, "", equation),
            (None, "最终结果"),
            (None, "7、下一题"),
        ])
        MODULE.separate_answers(source, output, force=False)
        texts = [text for _, text in text_and_styles(output)]
        answer_index = texts.index("第一章 题目答案")
        self.assertNotIn("【万人教育解析】", texts[:answer_index])
        self.assertIn("6.【万人教育解析】", texts[answer_index:])
        self.assertIn("最终结果", texts[answer_index:])
        with zipfile.ZipFile(output) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        self.assertEqual(1, len(root.findall(".//m:t[.='11010.1=26.5']", NS)))

    def test_nonconsecutive_numbered_answer_table_is_left_for_manual_review(self) -> None:
        result = self.separate([
            ("h1", "第一章 题目"),
            ("h2", "二、填空题"),
            (None, "1、第一题______"),
            (None, "2、第二题______"),
            (None, "【万人教育解析】"),
            (None, "1、答案甲"),
            (None, "3、答案丙"),
        ])
        texts = [text for _, text in result]
        self.assertIn("1、答案甲", texts)
        self.assertIn("3、答案丙", texts)
        self.assertIn("需人工复核", texts)

    def test_answer_heading_advances_only_the_major_part_of_nested_numbering(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "4.本章例题"),
            ("h3", "4.1.1选择题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
        ])
        texts = [text for _, text in result]
        self.assertIn("5.本章例题答案", texts)
        self.assertIn("5.1.1选择题", texts)
        self.assertNotIn("6.1.1选择题", texts)

    def test_numbered_question_text_containing_question_keyword_is_not_a_heading(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "本章例题"),
            (None, "1. 请说明这个题目的关键步骤。"),
            (None, "【万人教育解析】关键步骤说明"),
            ("h1", "第二章"),
        ])
        texts = [text for _, text in result]
        self.assertEqual(texts.count("本章例题答案"), 1)
        self.assertIn("1.【万人教育解析】关键步骤说明", texts)

    def test_rich_answer_and_continuation_are_moved_without_loss(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        formula_object = "<w:r><w:object><w:objectEmbed/></w:object></w:r>"
        make_document(source, [
            ("h1", "第二讲 三种调速方法简介（题目）"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】公式见后", formula_object),
            (None, "这是紧随的解释段。"),
            (None, "2. 第二题"),
            (None, "【万人教育解析】第二题答案"),
        ])
        MODULE.separate_answers(source, output, force=False)
        texts = [text for _, text in text_and_styles(output)]
        answer_title = "第二讲 三种调速方法简介（题目答案）"
        answer_index = texts.index(answer_title)
        self.assertNotIn("【万人教育解析】公式见后", texts[:answer_index])
        self.assertNotIn("这是紧随的解释段。", texts[:answer_index])
        self.assertIn("1.【万人教育解析】公式见后", texts[answer_index:])
        self.assertIn("这是紧随的解释段。", texts[answer_index:])
        self.assertIn(("1.【万人教育解析】公式见后", True), paragraphs_with_objects(output))

    def test_keyword_heading_and_ambiguous_answers_are_not_silently_moved(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "课后练习题"),
            (None, "1. 有歧义的题目"),
            (None, "【万人教育解析】答案甲"),
            (None, "【万人教育解析】答案乙"),
        ])
        texts = [text for _, text in result]
        self.assertIn("需人工复核", texts)
        self.assertNotIn("课后练习题答案", texts)
        self.assertIn("【万人教育解析】答案甲", texts)
        self.assertIn("【万人教育解析】答案乙", texts)

    def test_rerunning_output_does_not_create_a_second_answer_area(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        first = Path(directory.name) / "first.docx"
        second = Path(directory.name) / "second.docx"
        make_document(source, [
            ("h1", "第一章"),
            ("h2", "本章例题"),
            ("h3", "选择题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
        ])
        first_issue_count, _ = MODULE.separate_answers(source, first, force=False)
        second_issue_count, _ = MODULE.separate_answers(first, second, force=False)
        texts = [text for _, text in text_and_styles(second)]
        self.assertEqual(texts.count("本章例题答案"), 1)
        self.assertLessEqual(second_issue_count, first_issue_count)
        self.assertLessEqual(texts.count("需人工复核"), 1)

    def test_font_availability_is_not_written_as_a_review_issue(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "本章例题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
        ])
        self.assertNotIn("需人工复核", [text for _, text in result])

    def test_legacy_font_warning_review_page_is_removed(self) -> None:
        result = self.separate([
            ("h1", "第一章"),
            ("h2", "本章例题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
            ("h1", "需人工复核"),
            (None, "未检测到华文新魏；输出已指定该字体，Word 可能会临时替代显示。"),
        ])
        texts = [text for _, text in result]
        self.assertNotIn("需人工复核", texts)
        self.assertFalse(any("未检测到华文新魏" in text for text in texts))

    def test_missing_toc_is_created_before_first_level_one_heading_and_formatted(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, [
            ("h1", "第一章"),
            ("h2", "本章例题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
        ])
        MODULE.separate_answers(source, output, force=False)

        document = package_xml(output, "word/document.xml")
        body = document.find("w:body", NS)
        self.assertIsNotNone(body)
        toc = [
            node
            for node in document.findall(".//w:sdt", NS)
            if node.find("./w:sdtPr/w:docPartObj/w:docPartGallery[@w:val='Table of Contents']", NS)
            is not None
        ]
        self.assertEqual(len(toc), 1)
        first_heading = next(
            paragraph
            for paragraph in body.findall("w:p", NS)
            if "第一章" == "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))
        )
        self.assertLess(list(body).index(toc[0]), list(body).index(first_heading))
        self.assertIsNotNone(toc[0].find(".//w:br[@w:type='page']", NS))
        self.assertIn('TOC \\o "1-3" \\h \\z \\u', "".join(node.text or "" for node in toc[0].findall(".//w:instrText", NS)))

        title = toc[0].find("./w:sdtContent/w:p", NS)
        self.assertEqual("目录", "".join(node.text or "" for node in title.findall(".//w:t", NS)))
        self.assertEqual("center", title.find("./w:pPr/w:jc", NS).get(VAL))
        title_fonts = title.find(".//w:rPr/w:rFonts", NS)
        self.assertEqual("华文新魏", title_fonts.get(qn("eastAsia")))
        self.assertEqual("96", title.find(".//w:rPr/w:sz", NS).get(VAL))
        self.assertEqual("96", title.find("./w:pPr/w:rPr/w:sz", NS).get(VAL))
        self.assertEqual("0", title.find(".//w:rPr/w:b", NS).get(VAL))

        styles = package_xml(output, "word/styles.xml")
        toc1 = styles.find("./w:style[@w:styleId='TOC1']", NS)
        toc1_fonts = toc1.find("./w:rPr/w:rFonts", NS)
        self.assertEqual("宋体", toc1_fonts.get(qn("eastAsia")))
        self.assertEqual("Times New Roman", toc1_fonts.get(qn("ascii")))
        self.assertEqual("0", toc1.find("./w:rPr/w:b", NS).get(VAL))

        field_begin = toc[0].find(".//w:fldChar[@w:fldCharType='begin']", NS)
        self.assertEqual("true", field_begin.get(qn("dirty")))

    def test_existing_toc_is_preserved_once_and_reformatted(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        source = Path(directory.name) / "source.docx"
        output = Path(directory.name) / "output.docx"
        make_document(source, [
            ("h1", "第一章"),
            ("h2", "本章例题"),
            (None, "1. 第一题"),
            (None, "【万人教育解析】答案"),
        ], include_toc=True)
        MODULE.separate_answers(source, output, force=False)

        toc = toc_nodes(output)
        self.assertEqual(len(toc), 1)
        title = toc[0].find("./w:sdtContent/w:p", NS)
        self.assertEqual("96", title.find(".//w:rPr/w:sz", NS).get(VAL))
        self.assertEqual("center", title.find("./w:pPr/w:jc", NS).get(VAL))


if __name__ == "__main__":
    unittest.main()
