#!/usr/bin/env python3
"""Separate immediately following 万人教育解析 answers from a .docx question bank."""

from __future__ import annotations

import argparse
import copy
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS, "m": M_NS}
P = f"{{{W_NS}}}p"
PPR = f"{{{W_NS}}}pPr"
R = f"{{{W_NS}}}r"
RPR = f"{{{W_NS}}}rPr"
T = f"{{{W_NS}}}t"
MT = f"{{{M_NS}}}t"
RFONTS = f"{{{W_NS}}}rFonts"
VAL = f"{{{W_NS}}}val"

for prefix, uri in {
    "w": W_NS,
    "m": M_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/word16sdtdh",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/word16sdtfl",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}.items():
    ET.register_namespace(prefix, uri)

ANSWER_MARKER = "【万人教育解析】"
LABEL = "万人教育解析"
LEGACY_FONT_WARNING = "未检测到华文新魏"
# A dotted number may be followed by a numeric acronym such as `6.8086 CPU`.
# Reject a decimal-looking value whose fractional part ends before punctuation,
# an operator, or the end of the paragraph, e.g. `11010.1=26.5`.
QUESTION_RE = re.compile(
    r"^\s*([1-9]\d*)\s*(?:[．]|、(?!\s*\d+\s*[KMG]?\s*\*)|\.(?!\s*(?:"
    r"\d(?=\s)|\d{1,3}(?=\s*(?:[,，。;；:：=+*/×÷]|$))|"
    r"\d+(?:\.\d+)?\s*(?:mV|V|mA|A|Hz|kHz|MHz|H)\b)))"
)
# Some teaching-style question banks use visible labels rather than a leading
# Arabic number.  These labels are deliberately narrow: notes, tips, analysis,
# and worked-solution labels such as `【注】` or `【做题步骤】` are not questions.
# `【题n】` retains its source number; an unnumbered `【例】` receives a local,
# manual display number only after a marker-bounded answer is proven.
BRACKETED_QUESTION_RE = re.compile(r"^\s*【(?:例|题\s*([1-9]\d*))】")
EMBEDDED_BRACKETED_QUESTION_RE = re.compile(r"【(?:例|题\s*[1-9]\d*)】")
EMBEDDED_NEXT_QUESTION_RE = re.compile(
    r"(?<!\d)(\d{1,2})\s*[．、.]\s*(?=[\u4e00-\u9fffA-Za-z（(]|\d{2,}(?=\s*[A-Za-z\u4e00-\u9fff]))"
)
SUBQUESTION_ANSWER_PREFIX_BODY = r"题\s*(?:[（(][^）)]+[）)])+\s*解\s*[：:]"
ANSWER_PREFIX_BODY = (
    r"(?:【万人教育解析】|"
    + SUBQUESTION_ANSWER_PREFIX_BODY
    + r"|(?:题目分析|答案|解答|解析|分析|答|解)\s*[：:])\s*"
)
ANSWER_PREFIX_RE = re.compile(r"^\s*" + ANSWER_PREFIX_BODY)
EXPLANATION_PREFIX_RE = re.compile(r"^\s*(?:题目分析|解析|分析)\s*[：:]")
SUBQUESTION_ANSWER_PREFIX_RE = re.compile(r"^\s*" + SUBQUESTION_ANSWER_PREFIX_BODY)
SUBANSWER_PART_RE = re.compile(r"^\s*[（(]\s*\d+\s*[）)]")
# Inside a question line, aliases must be independent labels.  A bare `答：`
# occurs naturally at the end of words such as `请回答：`, which is not an
# answer boundary.  The canonical bracket marker remains valid anywhere.
INLINE_ANSWER_PREFIX_RE = re.compile(
    r"【万人教育解析】\s*|(?<![\u4e00-\u9fffA-Za-z0-9])"
    r"(?:题目分析|答案|解答|解析|分析|答|解)\s*[：:]\s*"
)
RANGE_ANSWER_RE = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*[：:]\s*([A-Za-z]+)\s*$")
OPTION_LINE_RE = re.compile(r"^\s*(?:[A-DＡ-Ｄ]|[①②③④⑤⑥⑦⑧⑨⑩])\s*[.．、:：]")
MATERIAL_LINE_RE = re.compile(r"^\s*(?:规格|控制信号|格式如下|主要包括)\s*[：:]?")
UNNUMBERED_QUESTION_CUE_RE = re.compile(r"(?:试编写|请编写)")
SECTION_REFERENCE_HEADING_RE = re.compile(
    r"^\s*\d+\s*[\.．]\s*\d+(?:\s*[、,，]\s*\d+\s*[\.．]\s*\d+)+\s*[^。；：？！()（）]*$"
)
HEADING_NUMBER_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百千万0-9]+[章节篇]\s*|\d+(?:\s*[\.．、]\s*\d+)*\s*[\.．、]?\s*)"
)
HEADING_NAME_RE = re.compile(r"(?:heading|标题)\s*([1-9]\d*)", re.IGNORECASE)
QUESTION_KEYWORDS = (
    "例题",
    "习题",
    "练习题",
    "题目",
    "测试题",
    "自测题",
    "思考题",
    "选择题",
    "单选题",
    "多选题",
    "填空题",
    "判断题",
    "简答题",
    "论述题",
    "综合题",
    "计算题",
    "案例题",
)
TYPE_HEADING_RE = re.compile(r"^(?:选择|单选|多选|填空|判断|简答|论述|综合|计算|案例|分析)题$")
TYPE_HEADING_PREFIX_RE = re.compile(r"^(?:[一二三四五六七八九十百千万0-9]+\s*[、.．]\s*)")
EMBEDDED_TYPE_HEADING_RE = re.compile(
    r"(?:[一二三四五六七八九十百千万]+\s*[、.．]\s*)?(?:选择|单选|多选|填空|判断|简答|论述|综合|计算|案例|分析)题"
)
GROUP_HEADING_SUFFIXES = ("例题", "习题", "练习题", "测试题", "自测题", "思考题", "题目")
LEADING_MAJOR_NUMBER_RE = re.compile(r"^(\s*)(\d+)(?=\s*[\.．])")


@dataclass(frozen=True)
class Heading:
    paragraph: ET.Element
    index: int
    level: int
    text: str
    normalized: str
    is_question: bool
    is_type: bool
    is_answer: bool


@dataclass(frozen=True)
class AnswerScope:
    title: Heading
    answer_title: str


def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter() if node.tag in {T, MT})


def word_paragraph_text(paragraph: ET.Element) -> str:
    """Return only ordinary Word text for distinguishing question numbers from formulas."""
    return "".join(node.text or "" for node in paragraph.iter(T))


def question_start_match(text: str) -> re.Match[str] | None:
    """Return a recognized numbered or bracket-labelled question prefix."""
    return QUESTION_RE.match(text) or BRACKETED_QUESTION_RE.match(text)


def question_number_from_text(text: str) -> str | None:
    """Return the explicit source number, if the question form provides one."""
    numbered = QUESTION_RE.match(text)
    if numbered is not None:
        return numbered.group(1)
    labelled = BRACKETED_QUESTION_RE.match(text)
    return labelled.group(1) if labelled is not None else None


def direct_paragraphs(body: ET.Element) -> list[ET.Element]:
    return [child for child in list(body) if child.tag == P]


OBJECT_PAYLOAD_TAGS = {
    qn("object"),
    qn("drawing"),
    qn("pict"),
    f"{{{M_NS}}}oMath",
    f"{{{M_NS}}}oMathPara",
}


def has_indivisible_payload(element: ET.Element) -> bool:
    """Return whether a direct paragraph child must belong to one fragment only."""
    return element.tag in OBJECT_PAYLOAD_TAGS or any(
        descendant.tag in OBJECT_PAYLOAD_TAGS for descendant in element.iter()
    )


def paragraph_text_fragment(paragraph: ET.Element, start: int, end: int) -> ET.Element:
    """Copy one visible-text interval without duplicating drawings/OLE/formulas."""
    fragment = ET.Element(P, paragraph.attrib)
    ppr = paragraph.find(PPR)
    if ppr is not None:
        fragment.append(copy.deepcopy(ppr))

    cursor = 0
    for child in list(paragraph):
        if child.tag == PPR:
            continue
        child_length = len(paragraph_text(child))
        child_end = cursor + child_length
        if has_indivisible_payload(child):
            # Payload-bearing XML is atomic. Its source position chooses one
            # fragment, so relationship-backed objects can never be cloned.
            owner_position = cursor
            if (start <= owner_position < end) or (
                end == len(paragraph_text(paragraph)) and owner_position == end
            ):
                fragment.append(copy.deepcopy(child))
        elif child_length:
            overlap_start = max(start, cursor)
            overlap_end = min(end, child_end)
            if overlap_start < overlap_end:
                copied = copy.deepcopy(child)
                retain_leading_text(copied, overlap_end - cursor)
                strip_leading_text(copied, overlap_start - cursor)
                fragment.append(copied)
        elif start <= cursor < end or (end == len(paragraph_text(paragraph)) and cursor == end):
            fragment.append(copy.deepcopy(child))
        cursor = child_end
    return fragment


def replace_paragraph_contents(paragraph: ET.Element, replacement: ET.Element) -> None:
    """Replace a paragraph's XML children while keeping its node identity."""
    paragraph.attrib.clear()
    paragraph.attrib.update(replacement.attrib)
    for child in list(paragraph):
        paragraph.remove(child)
    for child in list(replacement):
        paragraph.append(child)


def split_joined_answer_and_question_paragraphs(body: ET.Element) -> None:
    """Split marked/alias answers joined to the next question without cloning objects."""
    index = 0
    while index < len(direct_paragraphs(body)):
        paragraph = direct_paragraphs(body)[index]
        text = paragraph_text(paragraph)
        leading_type = EMBEDDED_TYPE_HEADING_RE.search(text)
        if leading_type is not None and text[: leading_type.start()].strip():
            leading_type = None
        type_question_boundary: int | None = None
        if leading_type is not None:
            following_question = EMBEDDED_NEXT_QUESTION_RE.search(text, leading_type.end())
            if (
                following_question is not None
                and not text[leading_type.end() : following_question.start(1)].strip()
            ):
                type_question_boundary = following_question.start(1)
        embedded_label_boundaries = [
            match.start()
            for match in EMBEDDED_BRACKETED_QUESTION_RE.finditer(text)
            if match.start() > 0
            and text[: match.start()].rstrip()
            and text[: match.start()].rstrip()[-1] in "。！？!?；;.）)"
        ]

        if (
            ANSWER_MARKER not in text
            and INLINE_ANSWER_PREFIX_RE.search(text) is None
            and type_question_boundary is None
            and not embedded_label_boundaries
        ):
            index += 1
            continue
        first_marker = INLINE_ANSWER_PREFIX_RE.search(text)

        def follows_answer_sentence(position: int) -> bool:
            prefix = text[:position].rstrip()
            return bool(prefix) and prefix[-1] in "。！？!?；;.）)"

        # A marker may be attached to the answer that follows the embedded
        # question (`…4.题干【标记】答案`).  Split every clear boundary in that
        # marked paragraph in one pass, including a Chinese type heading.
        boundaries = [*([type_question_boundary] if type_question_boundary is not None else [])]
        boundaries.extend(embedded_label_boundaries)
        boundaries.extend(
            match.start(1)
            for match in EMBEDDED_NEXT_QUESTION_RE.finditer(text)
            if follows_answer_sentence(match.start(1))
        )
        boundaries.extend(
            match.start()
            for match in EMBEDDED_TYPE_HEADING_RE.finditer(text)
            if follows_answer_sentence(match.start())
            and not re.search(r"\d+\s*[.．]\s*\d+\s*$", text[: match.start()].strip())
        )
        boundaries = sorted(set(boundaries))
        if not boundaries:
            index += 1
            continue
        starts = [0, *boundaries]
        ends = [*boundaries, len(text)]
        parts: list[ET.Element] = []
        for start, end in zip(starts, ends):
            parts.append(paragraph_text_fragment(paragraph, start, end))
        body_index = list(body).index(paragraph)
        body.remove(paragraph)
        # `index` counts only paragraphs, whereas the body also contains
        # tables and section properties.  Insert relative to the actual XML
        # child so an embedded split never jumps across a table/chapter.
        for offset, part in enumerate(parts):
            body.insert(body_index + offset, part)
        # Revisit the newly inserted fragments. A source paragraph may contain
        # several generations of joins (answer → question → answer → type
        # heading → question), and the second boundary becomes visible only
        # after the first split.


def prepend_source_question_number(paragraph: ET.Element, question_number: str) -> None:
    """Restore an omitted source question number without inheriting label font."""
    prefix = ET.Element(R)
    template = paragraph.find(R)
    if template is not None and template.find(RPR) is not None:
        prefix.append(copy.deepcopy(template.find(RPR)))
    ET.SubElement(prefix, T).text = f"{question_number}、"
    ppr = paragraph.find(PPR)
    insert_index = list(paragraph).index(ppr) + 1 if ppr is not None else 0
    paragraph.insert(insert_index, prefix)


def restore_omitted_question_numbers(body: ET.Element, headings: list[Heading]) -> None:
    """Recover only structurally certain unnumbered questions.

    Some source chapters omit `4、` between `答：…` and the next fill-in
    prompt, or omit the first number of a newly numbered subsection.  Restore
    a number only if the prompt has blanks/a question cue and a recognized
    answer marker follows before the next numbered question.
    """
    paragraphs = direct_paragraphs(body)
    texts = [paragraph_text(paragraph).strip() for paragraph in paragraphs]
    heading_indices = {heading.index for heading in headings}
    for index, paragraph in enumerate(paragraphs):
        text = texts[index]
        if not text or question_start_match(word_paragraph_text(paragraph)) or answer_prefix_match(text):
            continue
        question_like = "___" in text or UNNUMBERED_QUESTION_CUE_RE.search(text) is not None
        if not question_like:
            continue
        # A bare underline inside assembly/source code is content, never an
        # independently recoverable question.
        if re.fullmatch(r"[_＿\s]+", text):
            continue
        previous_nonempty = next((texts[position] for position in range(index - 1, -1, -1) if texts[position]), "")
        # A lone blank inside an assembly listing is usually a sub-step, not a
        # new numbered question.  Fill-in prompts are recoverable only after
        # the preceding numbered answer; cue-based program questions remain
        # valid at the start of a subsection.
        if "___" in text and UNNUMBERED_QUESTION_CUE_RE.search(text) is None and not answer_prefix_match(previous_nonempty):
            continue
        following = range(index + 1, min(len(paragraphs), index + 7))
        answer_position = next((position for position in following if answer_prefix_match(texts[position])), None)
        next_numbered = next(
            (position for position in following if question_start_match(word_paragraph_text(paragraphs[position]))),
            None,
        )
        next_heading = next((position for position in following if position in heading_indices), None)
        if (
            answer_position is None
            or (next_numbered is not None and next_numbered < answer_position)
            or (next_heading is not None and next_heading < answer_position)
        ):
            continue
        previous_numbers: list[int] = []
        for position in range(index - 1, -1, -1):
            if position in heading_indices:
                previous_numbers.clear()
                break
            previous_match = QUESTION_RE.match(word_paragraph_text(paragraphs[position]))
            if previous_match is not None:
                previous_numbers.append(int(previous_match.group(1)))
                break
        number = previous_numbers[0] + 1 if previous_numbers else 1
        prepend_source_question_number(paragraph, str(number))


def remove_legacy_font_warning_review(body: ET.Element) -> None:
    """Remove only the obsolete font-availability review page written by old versions."""
    paragraphs = direct_paragraphs(body)
    warning_indexes = [
        index for index, paragraph in enumerate(paragraphs) if LEGACY_FONT_WARNING in paragraph_text(paragraph)
    ]
    if not warning_indexes:
        return
    for index in reversed(warning_indexes):
        body.remove(paragraphs[index])

    remaining = direct_paragraphs(body)
    for index, paragraph in reversed(list(enumerate(remaining))):
        if paragraph_text(paragraph).strip() != "需人工复核":
            continue
        if not any(paragraph_text(after).strip() for after in remaining[index + 1 :]):
            body.remove(paragraph)


def normalize_heading_text(text: str) -> str:
    return HEADING_NUMBER_RE.sub("", text).strip()


def is_answer_heading(text: str) -> bool:
    normalized = normalize_heading_text(text)
    return (
        normalized in {"本章习题答案", "本节习题答案", "本章例题答案"}
        or "题目答案" in normalized
        or normalized.endswith(("例题答案", "习题答案"))
    )


def answer_title_for(source_title: str) -> str:
    title = source_title.strip()
    parenthesized = re.sub(r"（\s*题目\s*）", "（题目答案）", title)
    parenthesized = re.sub(r"\(\s*题目\s*\)", "(题目答案)", parenthesized)
    if parenthesized != title:
        return parenthesized
    if "题目" in title:
        return title[: title.rfind("题目")] + "题目答案" + title[title.rfind("题目") + len("题目") :]
    if title.endswith(("例题", "习题")):
        return title + "答案"
    return f"{title}（题目答案）"


def advance_heading_major_number(text: str) -> str:
    """Advance only a dotted heading's first numeric component."""
    match = LEADING_MAJOR_NUMBER_RE.match(text)
    if match is None:
        return text
    return f"{match.group(1)}{int(match.group(2)) + 1}{text[match.end(2):]}"


def is_question_heading(text: str) -> bool:
    normalized = normalize_heading_text(text)
    return not is_answer_heading(text) and any(keyword in normalized for keyword in QUESTION_KEYWORDS)


def is_type_heading(text: str) -> bool:
    normalized = TYPE_HEADING_PREFIX_RE.sub("", normalize_heading_text(text))
    return bool(TYPE_HEADING_RE.fullmatch(normalized))


def style_outline_levels(styles_xml: bytes | None) -> dict[str, int]:
    if not styles_xml:
        return {}
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        return {}

    styles = {
        style.get(qn("styleId")): style
        for style in root.findall("w:style", NS)
        if style.get(qn("styleId"))
    }
    resolved: dict[str, int | None] = {}

    def resolve(style_id: str, seen: set[str]) -> int | None:
        if style_id in resolved:
            return resolved[style_id]
        if style_id in seen or style_id not in styles:
            return None
        style = styles[style_id]
        seen = seen | {style_id}
        outline = style.find("./w:pPr/w:outlineLvl", NS)
        if outline is not None and outline.get(VAL) is not None:
            resolved[style_id] = int(outline.get(VAL)) + 1
            return resolved[style_id]
        name = style.find("./w:name", NS)
        name_value = name.get(VAL, "") if name is not None else ""
        name_match = HEADING_NAME_RE.search(name_value)
        if name_match:
            resolved[style_id] = int(name_match.group(1))
            return resolved[style_id]
        based_on = style.find("./w:basedOn", NS)
        based_on_id = based_on.get(VAL) if based_on is not None else None
        resolved[style_id] = resolve(based_on_id, seen) if based_on_id else None
        return resolved[style_id]

    return {style_id: level for style_id in styles if (level := resolve(style_id, set())) is not None}


def fallback_heading_level(text: str) -> int | None:
    normalized = normalize_heading_text(text)
    if SECTION_REFERENCE_HEADING_RE.fullmatch(text):
        return 2
    if is_type_heading(text):
        return 3
    if not is_question_heading(text):
        return None
    # A numbered question sentence can contain words such as “题目”; without a
    # Word heading style it must still look like a compact title before using
    # numeric hierarchy as a fallback.
    if (
        "本章" not in normalized
        and "本节" not in normalized
        and not is_type_heading(text)
        and (
            len(normalized) > 30
            or any(mark in normalized for mark in "。；：？！（）()")
            or not normalized.endswith(GROUP_HEADING_SUFFIXES)
        )
    ):
        return None
    number_match = re.match(r"^\s*(\d+(?:\s*[\.．]\s*\d+)*)", text)
    if number_match:
        components = re.split(r"\s*[\.．]\s*", number_match.group(1))
        return len(components) + 1
    if normalized.startswith("第") and any(unit in normalized for unit in ("章", "篇")):
        return 1
    return None


def paragraph_heading_level(paragraph: ET.Element, style_levels: dict[str, int]) -> int | None:
    ppr = paragraph.find(PPR)
    if ppr is None:
        return fallback_heading_level(paragraph_text(paragraph))
    outline = ppr.find(qn("outlineLvl"))
    if outline is not None and outline.get(VAL) is not None:
        return int(outline.get(VAL)) + 1
    style = ppr.find(qn("pStyle"))
    style_id = style.get(VAL) if style is not None else None
    return style_levels.get(style_id) if style_id else fallback_heading_level(paragraph_text(paragraph))


def collect_headings(paragraphs: list[ET.Element], style_levels: dict[str, int]) -> list[Heading]:
    headings: list[Heading] = []
    for index, paragraph in enumerate(paragraphs):
        level = paragraph_heading_level(paragraph, style_levels)
        if level is None:
            continue
        text = paragraph_text(paragraph)
        headings.append(
            Heading(
                paragraph=paragraph,
                index=index,
                level=level,
                text=text,
                normalized=normalize_heading_text(text),
                is_question=is_question_heading(text),
                is_type=is_type_heading(text),
                is_answer=is_answer_heading(text),
            )
        )
    return headings


def ancestors_of(heading: Heading, headings: list[Heading]) -> list[Heading]:
    ancestors: list[Heading] = []
    ceiling = heading.level
    for previous in reversed(headings[: headings.index(heading)]):
        if previous.level < ceiling:
            ancestors.append(previous)
            ceiling = previous.level
    return ancestors


def collect_answer_scopes(headings: list[Heading], paragraphs: list[ET.Element]) -> list[AnswerScope]:
    """Find explicit question groups and chapters evidenced by marked Q&A pairs.

    Some source banks use ordinary chapter/section names (for example,
    ``第二章 8086 CPU 的结构与功能``) rather than a heading containing “题目” or
    “习题”.  Such a chapter is safe to process only after a numbered question
    has a recognized answer marker before the next numbered question or heading.
    """
    scopes: dict[ET.Element, AnswerScope] = {}
    for heading in headings:
        if not heading.is_question or heading.is_answer:
            continue
        ancestors = ancestors_of(heading, headings)
        # Generated answer areas preserve type headings for readability. Their
        # descendants are output, not a fresh source question group on rerun.
        if any(ancestor.is_answer for ancestor in ancestors):
            continue
        question_ancestor = next(
            (ancestor for ancestor in ancestors if ancestor.is_question and not ancestor.is_type and not ancestor.is_answer),
            None,
        )

        if question_ancestor is not None:
            owner = question_ancestor
        elif heading.is_type:
            parent = ancestors[0] if ancestors else heading
            owner = parent
        else:
            owner = heading
        scopes[owner.paragraph] = AnswerScope(owner, advance_heading_major_number(answer_title_for(owner.text)))

    heading_indices = {heading.index for heading in headings}
    question_indices = [
        index
        for index, paragraph in enumerate(paragraphs)
        if index not in heading_indices and question_start_match(word_paragraph_text(paragraph))
    ]
    labelled_question_indices = [
        index
        for index, paragraph in enumerate(paragraphs)
        if index not in heading_indices and BRACKETED_QUESTION_RE.match(word_paragraph_text(paragraph))
    ]
    question_boundaries = set(question_indices) | heading_indices
    all_texts = [paragraph_text(paragraph) for paragraph in paragraphs]
    explicit_scope_ranges = [
        (scope.title.index, scope_end(scope, headings, paragraphs)) for scope in scopes.values()
    ]

    def evidence_is_sufficient(pair_evidence: list[bool]) -> bool:
        paired_count = sum(pair_evidence)
        return (
            len(pair_evidence) >= 2
            and paired_count >= 2
            and paired_count / len(pair_evidence) >= 0.8
        )

    def active_heading_chain(question_index: int) -> list[Heading]:
        chain: list[Heading] = []
        ceiling = float("inf")
        for heading in reversed([heading for heading in headings if heading.index < question_index]):
            if heading.level < ceiling:
                chain.append(heading)
                ceiling = heading.level
        return list(reversed(chain))

    # A solved-example guide often uses `【例】` / `【题n】` for its questions,
    # then numbers steps inside the solution (for example, `1. 计算…`).  Build
    # chapter evidence from the explicit labels alone so solution steps cannot
    # dilute the safety ratio or create a narrower, accidental section scope.
    labelled_boundaries = set(labelled_question_indices) | heading_indices
    labelled_evidence: dict[ET.Element, tuple[Heading, list[bool]]] = {}
    for question_index in labelled_question_indices:
        if any(start < question_index < end for start, end in explicit_scope_ranges):
            continue
        preceding_headings = [heading for heading in headings if heading.index < question_index]
        if not preceding_headings:
            continue
        nearest_heading = preceding_headings[-1]
        nearest_chapter = next(
            (heading for heading in reversed(preceding_headings) if heading.level == 1),
            None,
        )
        if nearest_heading.is_answer or (nearest_chapter is not None and nearest_chapter.is_answer):
            continue
        boundary = next(
            (index for index in range(question_index + 1, len(paragraphs)) if index in labelled_boundaries),
            len(paragraphs),
        )
        question_match = BRACKETED_QUESTION_RE.match(word_paragraph_text(paragraphs[question_index]))
        inline_match = inline_answer_prefix_match(all_texts[question_index], question_match)
        candidate_positions = [
            index
            for index in range(question_index + 1, boundary)
            if is_answer_candidate(all_texts[index])
        ]
        is_safe_pair = (
            inline_match is not None
            or candidate_positions_are_a_single_answer_block(
                candidate_positions,
                all_texts,
                allow_repeated_canonical_subanswers=True,
            )
        )
        owner = nearest_chapter or nearest_heading
        labelled_evidence.setdefault(owner.paragraph, (owner, []))[1].append(is_safe_pair)

    accepted_chapters: set[ET.Element] = set()
    for owner, pair_evidence in labelled_evidence.values():
        if not evidence_is_sufficient(pair_evidence):
            continue
        scopes.setdefault(owner.paragraph, AnswerScope(owner, answer_title_for(owner.text)))
        accepted_chapters.add(owner.paragraph)

    implicit_evidence: dict[ET.Element, tuple[Heading, list[bool]]] = {}
    section_evidence: dict[ET.Element, tuple[Heading, Heading, list[bool]]] = {}
    for question_index in question_indices:
        if any(start < question_index < end for start, end in explicit_scope_ranges):
            continue
        preceding_headings = [heading for heading in headings if heading.index < question_index]
        if not preceding_headings:
            continue
        nearest_heading = preceding_headings[-1]
        nearest_chapter = next(
            (heading for heading in reversed(preceding_headings) if heading.level == 1),
            None,
        )
        if nearest_heading.is_answer or (nearest_chapter is not None and nearest_chapter.is_answer):
            continue
        if nearest_chapter is not None and nearest_chapter.paragraph in accepted_chapters:
            continue
        boundary = next(
            (index for index in range(question_index + 1, len(paragraphs)) if index in question_boundaries),
            len(paragraphs),
        )
        # A chapter-wide area keeps all intermediate section headings together,
        # preventing overlapping answer areas for nested, unlabelled sections.
        owner = nearest_chapter or nearest_heading
        question_match = question_start_match(word_paragraph_text(paragraphs[question_index]))
        inline_match = inline_answer_prefix_match(all_texts[question_index], question_match)
        candidate_positions = [
            index
            for index in range(question_index + 1, boundary)
            if is_answer_candidate(all_texts[index])
        ]
        is_safe_pair = (
            inline_match is not None
            or candidate_positions_are_a_single_answer_block(candidate_positions, all_texts)
        )
        implicit_evidence.setdefault(owner.paragraph, (owner, []))[1].append(
            is_safe_pair
        )
        if nearest_chapter is not None:
            chain = active_heading_chain(question_index)
            try:
                chapter_position = chain.index(nearest_chapter)
            except ValueError:
                continue
            if chapter_position + 1 < len(chain):
                section_owner = chain[chapter_position + 1]
                if not section_owner.is_answer:
                    section_evidence.setdefault(
                        section_owner.paragraph,
                        (section_owner, nearest_chapter, []),
                    )[2].append(
                        is_safe_pair
                    )

    for owner, pair_evidence in implicit_evidence.values():
        # Require a substantial, non-trivial consistent sample. This permits
        # ordinary chapter names while keeping mixed notes/solutions out of an
        # inferred scope, where automatic movement would be unsafe.
        if not evidence_is_sufficient(pair_evidence):
            continue
        scopes.setdefault(
            owner.paragraph,
            AnswerScope(owner, answer_title_for(owner.text)),
        )
        accepted_chapters.add(owner.paragraph)
    # A mixed chapter can still have a clean, independently delimited section
    # (for example, 12.1 before formula-heavy 12.2).  Infer that direct section
    # only when the chapter itself did not qualify, avoiding overlapping areas.
    for owner, chapter, pair_evidence in section_evidence.values():
        if chapter.paragraph in accepted_chapters or not evidence_is_sufficient(pair_evidence):
            continue
        scopes.setdefault(
            owner.paragraph,
            AnswerScope(owner, answer_title_for(owner.text)),
        )
    return sorted(scopes.values(), key=lambda scope: scope.title.index)


def scope_end(scope: AnswerScope, headings: list[Heading], paragraphs: list[ET.Element]) -> int:
    following = next(
        (heading.index for heading in headings if heading.index > scope.title.index and heading.level <= scope.title.level),
        len(paragraphs),
    )
    return following


def first_run_properties(paragraph: ET.Element) -> ET.Element | None:
    run = paragraph.find(f".//{R}")
    if run is None:
        return None
    properties = run.find(RPR)
    return copy.deepcopy(properties) if properties is not None else None


def set_text(paragraph: ET.Element, text: str, run_properties: ET.Element | None = None) -> None:
    for child in list(paragraph):
        if child.tag != PPR:
            paragraph.remove(child)
    run = ET.SubElement(paragraph, R)
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    text_node = ET.SubElement(run, T)
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    text_node.text = text


def clone_text_paragraph(template: ET.Element, text: str) -> ET.Element:
    paragraph = copy.deepcopy(template)
    set_text(paragraph, text, first_run_properties(template))
    return paragraph


def clone_heading_with_advanced_number(template: ET.Element) -> ET.Element:
    paragraph = copy.deepcopy(template)
    for text_node in paragraph.iter(T):
        original = text_node.text or ""
        advanced = advance_heading_major_number(original)
        if advanced != original:
            text_node.text = advanced
            return paragraph
    full_text = paragraph_text(paragraph)
    advanced_full_text = advance_heading_major_number(full_text)
    if advanced_full_text != full_text:
        set_text(paragraph, advanced_full_text, first_run_properties(template))
    return paragraph


def make_run(parent: ET.Element, text: str, properties: ET.Element | None = None) -> None:
    if not text:
        return
    run = ET.SubElement(parent, R)
    if properties is not None:
        run.append(copy.deepcopy(properties))
    text_node = ET.SubElement(run, T)
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    text_node.text = text


def label_properties(template: ET.Element) -> ET.Element:
    properties = first_run_properties(template)
    if properties is None:
        properties = ET.Element(RPR)
    set_label_font(properties)
    return properties


def set_label_font(properties: ET.Element) -> None:
    fonts = properties.find(RFONTS)
    if fonts is None:
        fonts = ET.SubElement(properties, RFONTS)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(attr), "华文新魏")


def prepend_question_number(paragraph: ET.Element, question_number: str) -> None:
    """Add the display number without replacing runs, equations, or embedded objects."""
    prefix = ET.Element(R)
    properties = next(
        (
            copy.deepcopy(run.find(RPR))
            for run in paragraph.findall(R)
            if (run_text := "".join(node.text or "" for node in run.findall(T)))
            and ANSWER_MARKER not in run_text
            and run.find(RPR) is not None
        ),
        None,
    )
    if properties is not None:
        # Never let a source answer run styled as the label font leak into the
        # generated number.  The marker alone uses 华文新魏.
        fonts = properties.find(RFONTS)
        if fonts is not None and fonts.get(qn("eastAsia")) == "华文新魏":
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                fonts.set(qn(attr), "宋体")
        prefix.append(properties)
    text_node = ET.SubElement(prefix, T)
    text_node.text = f"{question_number}."
    ppr = paragraph.find(PPR)
    insert_index = list(paragraph).index(ppr) + 1 if ppr is not None else 0
    paragraph.insert(insert_index, prefix)


def answer_prefix_match(text: str) -> re.Match[str] | None:
    return ANSWER_PREFIX_RE.match(text)


def is_answer_candidate(text: str) -> bool:
    return answer_prefix_match(text) is not None or ANSWER_MARKER in text


def inline_answer_prefix_match(text: str, question_match: re.Match[str]) -> re.Match[str] | None:
    """Return a marker embedded after a question's display number, if any."""
    match = INLINE_ANSWER_PREFIX_RE.search(text, question_match.end())
    return match if match is not None and match.start() > question_match.end() else None


def candidate_positions_are_a_single_answer_block(
    positions: list[int],
    texts: list[str],
    *,
    allow_repeated_canonical_subanswers: bool = False,
) -> bool:
    """Allow an alias-led explanation followed by one final canonical answer."""
    if not positions:
        return False
    if len(positions) == 1:
        return True
    if (
        allow_repeated_canonical_subanswers
        and all(texts[position].lstrip().startswith(ANSWER_MARKER) for position in positions)
        and all(
            any(SUBANSWER_PART_RE.match(texts[between]) for between in range(previous + 1, current))
            for previous, current in zip(positions, positions[1:])
        )
    ):
        return True
    if all(SUBQUESTION_ANSWER_PREFIX_RE.match(texts[position]) for position in positions):
        return True
    if SUBQUESTION_ANSWER_PREFIX_RE.match(texts[positions[0]]) and all(
        SUBQUESTION_ANSWER_PREFIX_RE.match(texts[position])
        or EXPLANATION_PREFIX_RE.match(texts[position])
        for position in positions[1:]
    ):
        return True
    if EXPLANATION_PREFIX_RE.match(texts[positions[0]]):
        final_labels = [
            position
            for position in positions[1:]
            if not EXPLANATION_PREFIX_RE.match(texts[position])
        ]
        if len(final_labels) <= 1:
            return True
    first_match = answer_prefix_match(texts[positions[0]])
    # `【标记】简答` followed by `解析：说明` belongs to one answer block,
    # provided no second canonical marker starts another answer.
    if (
        texts[positions[0]].lstrip().startswith(ANSWER_MARKER)
        and all(not texts[position].lstrip().startswith(ANSWER_MARKER) for position in positions[1:])
    ):
        return True
    return (
        first_match is not None
        and not texts[positions[0]].lstrip().startswith(ANSWER_MARKER)
        and any(texts[position].lstrip().startswith(ANSWER_MARKER) for position in positions[1:])
    )


def markerless_answer_start(texts: list[str], question_index: int, end: int) -> int | None:
    """Recognize explanatory answer prose that follows a question without a label."""
    rows = [index for index in range(question_index + 1, end) if texts[index].strip()]
    if not rows or any(is_answer_candidate(texts[index]) for index in rows):
        return None
    first = texts[rows[0]].strip()
    if (
        question_start_match(first)
        or OPTION_LINE_RE.match(first)
        or any(MATERIAL_LINE_RE.match(texts[index].strip()) for index in rows)
    ):
        return None
    # A bare one-line sentence is too easy to confuse with ordinary prose.
    # Accept it only when it is substantively explanatory; short, unlabeled
    # Chinese text needs at least one continuation paragraph to be reliable.
    return rows[0] if len(rows) >= 2 or len(first) >= 30 else None


def strip_leading_text(paragraph: ET.Element, count: int) -> None:
    """Remove text from both Word runs and Office Math runs without touching objects."""
    remaining = count
    for text_node in paragraph.iter():
        if text_node.tag not in {T, MT}:
            continue
        value = text_node.text or ""
        if remaining >= len(value):
            text_node.text = ""
            remaining -= len(value)
        elif remaining:
            text_node.text = value[remaining:]
            remaining = 0


def remove_empty_math_containers(element: ET.Element) -> None:
    """Drop equation wrappers left empty after converting a formula label to Word text."""
    for child in list(element):
        remove_empty_math_containers(child)
        if child.tag in {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"} and not paragraph_text(child):
            element.remove(child)


def retain_leading_text(paragraph: ET.Element, count: int) -> None:
    """Keep only the first `count` visible text characters without dropping objects."""
    remaining = count
    for text_node in paragraph.iter():
        if text_node.tag not in {T, MT}:
            continue
        value = text_node.text or ""
        if remaining <= 0:
            text_node.text = ""
        elif remaining < len(value):
            text_node.text = value[:remaining]
            remaining = 0
        else:
            remaining -= len(value)


def insert_answer_marker(paragraph: ET.Element) -> None:
    label = ET.Element(R)
    label.append(label_properties(paragraph))
    text_node = ET.SubElement(label, T)
    text_node.text = ANSWER_MARKER
    ppr = paragraph.find(PPR)
    insert_index = list(paragraph).index(ppr) + 1 if ppr is not None else 0
    paragraph.insert(insert_index, label)


def clone_run_text(template: ET.Element, text: str, *, label: bool = False) -> ET.Element:
    run = ET.Element(R)
    properties = template.find(RPR)
    if properties is not None:
        properties = copy.deepcopy(properties)
    elif label:
        properties = ET.Element(RPR)
    if properties is not None:
        if label:
            set_label_font(properties)
        run.append(properties)
    text_node = ET.SubElement(run, T)
    if text.startswith(" ") or text.endswith(" "):
        text_node.set(f"{{{XML_NS}}}space", "preserve")
    text_node.text = text
    return run


def normalize_answer_marker_fonts(paragraph: ET.Element) -> None:
    """Put the bracket label in its own 华文新魏 Word run when possible."""
    for run in list(paragraph.findall(R)):
        text_nodes = run.findall(T)
        if len(text_nodes) != 1:
            continue
        text = text_nodes[0].text or ""
        marker_index = text.find(ANSWER_MARKER)
        if marker_index < 0:
            continue
        if text == ANSWER_MARKER:
            properties = ensure_rpr(run)
            set_label_font(properties)
            continue
        before = text[:marker_index]
        after = text[marker_index + len(ANSWER_MARKER) :]
        replacement = [
            *([clone_run_text(run, before)] if before else []),
            clone_run_text(run, ANSWER_MARKER, label=True),
            *([clone_run_text(run, after)] if after else []),
        ]
        index = list(paragraph).index(run)
        paragraph.remove(run)
        for offset, new_run in enumerate(replacement):
            paragraph.insert(index + offset, new_run)


def replace_leading_answer_prefix(paragraph: ET.Element) -> None:
    """Normalize a leading answer label while preserving the remaining rich runs."""
    match = answer_prefix_match(paragraph_text(paragraph))
    if match is None:
        return
    strip_leading_text(paragraph, match.end())
    remove_empty_math_containers(paragraph)
    insert_answer_marker(paragraph)


def strip_leading_answer_prefix(paragraph: ET.Element) -> None:
    """Remove a redundant label from a continuation within one answer block."""
    match = answer_prefix_match(paragraph_text(paragraph))
    if match is not None:
        strip_leading_text(paragraph, match.end())
        remove_empty_math_containers(paragraph)


def compact_answer_paragraph(template: ET.Element, answer: str) -> ET.Element:
    paragraph = copy.deepcopy(template)
    set_text(paragraph, f"{ANSWER_MARKER}{answer}", first_run_properties(template))
    return paragraph


def clone_answer_block(paragraphs: list[ET.Element], question_number: str) -> list[ET.Element]:
    copied = [copy.deepcopy(paragraph) for paragraph in paragraphs]
    replace_leading_answer_prefix(copied[0])
    for continuation in copied[1:]:
        strip_leading_answer_prefix(continuation)
    prepend_question_number(copied[0], question_number)
    return copied


def clone_inline_answer(paragraph: ET.Element, marker_match: re.Match[str]) -> ET.Element:
    """Copy the embedded answer tail while the source paragraph retains its question."""
    copied = copy.deepcopy(paragraph)
    strip_leading_text(copied, marker_match.start())
    replace_leading_answer_prefix(copied)
    return copied


def clone_markerless_answer_block(paragraphs: list[ET.Element]) -> list[ET.Element]:
    copied = [copy.deepcopy(paragraph) for paragraph in paragraphs]
    insert_answer_marker(copied[0])
    return copied


def clone_numbered_answer_row(paragraph: ET.Element) -> ET.Element:
    """Turn an answer-list row such as `1、答案` into a marked answer paragraph."""
    copied = copy.deepcopy(paragraph)
    text = paragraph_text(copied)
    match = QUESTION_RE.match(text)
    if match is None:
        raise ValueError("编号答案行缺少题号。")
    remainder = text[match.end() :]
    strip_leading_text(copied, match.end() + len(remainder) - len(remainder.lstrip()))
    insert_answer_marker(copied)
    return copied


def ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    return child


def ensure_ppr(paragraph: ET.Element) -> ET.Element:
    ppr = paragraph.find(PPR)
    if ppr is None:
        ppr = ET.Element(PPR)
        paragraph.insert(0, ppr)
    return ppr


def ensure_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find(RPR)
    if rpr is None:
        rpr = ET.Element(RPR)
        run.insert(0, rpr)
    return rpr


def set_font_format(properties: ET.Element, east_asia: str, latin: str, *, size: str | None = None) -> None:
    fonts = ensure_child(properties, RFONTS)
    fonts.set(qn("ascii"), latin)
    fonts.set(qn("hAnsi"), latin)
    fonts.set(qn("cs"), latin)
    fonts.set(qn("eastAsia"), east_asia)
    for tag in (qn("b"), qn("bCs")):
        ensure_child(properties, tag).set(VAL, "0")
    if size is not None:
        for tag in (qn("sz"), qn("szCs")):
            ensure_child(properties, tag).set(VAL, size)


def toc_field_instruction(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//w:instrText", NS))


def toc_containers(root: ET.Element) -> list[ET.Element]:
    controls = [
        node
        for node in root.findall(".//w:sdt", NS)
        if node.find("./w:sdtPr/w:docPartObj/w:docPartGallery[@w:val='Table of Contents']", NS)
        is not None
    ]
    body = root.find("w:body", NS)
    direct_fields = [
        paragraph
        for paragraph in (direct_paragraphs(body) if body is not None else [])
        if "TOC" in toc_field_instruction(paragraph)
    ]
    return controls + direct_fields


def make_toc_title_paragraph() -> ET.Element:
    paragraph = ET.Element(P)
    ppr = ET.SubElement(paragraph, PPR)
    ET.SubElement(ppr, qn("pStyle")).set(VAL, "TOC")
    ET.SubElement(ppr, qn("jc")).set(VAL, "center")
    run = ET.SubElement(paragraph, R)
    properties = ET.SubElement(run, RPR)
    set_font_format(properties, "华文新魏", "华文新魏", size="96")
    ET.SubElement(run, T).text = "目录"
    return paragraph


def make_toc_field_paragraph() -> ET.Element:
    paragraph = ET.Element(P)
    ppr = ET.SubElement(paragraph, PPR)
    ET.SubElement(ppr, qn("pStyle")).set(VAL, "TOC1")
    begin = ET.SubElement(paragraph, R)
    field_begin = ET.SubElement(begin, qn("fldChar"))
    field_begin.set(qn("fldCharType"), "begin")
    field_begin.set(qn("dirty"), "true")
    instruction = ET.SubElement(paragraph, R)
    ET.SubElement(instruction, qn("instrText")).text = ' TOC \\o "1-3" \\h \\z \\u '
    separate = ET.SubElement(paragraph, R)
    ET.SubElement(separate, qn("fldChar")).set(qn("fldCharType"), "separate")
    placeholder = ET.SubElement(paragraph, R)
    placeholder_properties = ET.SubElement(placeholder, RPR)
    set_font_format(placeholder_properties, "宋体", "Times New Roman")
    ET.SubElement(placeholder, T).text = "目录将在打开文档时更新。"
    end = ET.SubElement(paragraph, R)
    ET.SubElement(end, qn("fldChar")).set(qn("fldCharType"), "end")
    return paragraph


def make_page_break_paragraph() -> ET.Element:
    paragraph = ET.Element(P)
    run = ET.SubElement(paragraph, R)
    ET.SubElement(run, qn("br")).set(qn("type"), "page")
    return paragraph


def make_toc_content_control() -> ET.Element:
    sdt = ET.Element(qn("sdt"))
    properties = ET.SubElement(sdt, qn("sdtPr"))
    ET.SubElement(properties, qn("alias")).set(VAL, "目录")
    ET.SubElement(properties, qn("tag")).set(VAL, "TOC")
    doc_part = ET.SubElement(properties, qn("docPartObj"))
    ET.SubElement(doc_part, qn("docPartGallery")).set(VAL, "Table of Contents")
    content = ET.SubElement(sdt, qn("sdtContent"))
    content.append(make_toc_title_paragraph())
    content.append(make_toc_field_paragraph())
    content.append(make_page_break_paragraph())
    return sdt


def toc_title_paragraph(container: ET.Element) -> ET.Element | None:
    search_root = container.find("./w:sdtContent", NS) if container.tag == qn("sdt") else container
    if search_root is None:
        return None
    paragraphs = search_root.findall(".//w:p", NS)
    return next((paragraph for paragraph in paragraphs if paragraph_text(paragraph).strip() == "目录"), None)


def format_toc_title(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    ensure_child(ppr, qn("jc")).set(VAL, "center")
    set_font_format(ensure_child(ppr, RPR), "华文新魏", "华文新魏", size="96")
    runs = paragraph.findall(".//w:r", NS)
    if not runs:
        run = ET.SubElement(paragraph, R)
        ET.SubElement(run, T).text = "目录"
        runs = [run]
    for run in runs:
        set_font_format(ensure_rpr(run), "华文新魏", "华文新魏", size="96")


def toc_entry_paragraphs(container: ET.Element) -> list[ET.Element]:
    search_root = container.find("./w:sdtContent", NS) if container.tag == qn("sdt") else container
    if search_root is None:
        return []
    result: list[ET.Element] = []
    for paragraph in search_root.findall(".//w:p", NS):
        pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
        style_id = pstyle.get(VAL, "") if pstyle is not None else ""
        if re.fullmatch(r"TOC[1-9]", style_id, re.IGNORECASE):
            result.append(paragraph)
    return result


def format_toc_entry_paragraph(paragraph: ET.Element) -> None:
    for run in paragraph.findall(".//w:r", NS):
        set_font_format(ensure_rpr(run), "宋体", "Times New Roman")


def find_style(root: ET.Element, style_id: str) -> ET.Element | None:
    return root.find(f"./w:style[@w:styleId='{style_id}']", NS)


def ensure_style(root: ET.Element, style_id: str, name: str) -> ET.Element:
    style = find_style(root, style_id)
    if style is None:
        style = ET.SubElement(root, qn("style"))
        style.set(qn("type"), "paragraph")
        style.set(qn("styleId"), style_id)
        ET.SubElement(style, qn("name")).set(VAL, name)
    return style


def format_toc_styles(styles_root: ET.Element) -> None:
    title_style = ensure_style(styles_root, "TOC", "TOC Heading")
    title_ppr = ensure_child(title_style, qn("pPr"))
    ensure_child(title_ppr, qn("jc")).set(VAL, "center")
    set_font_format(ensure_child(title_style, RPR), "华文新魏", "华文新魏", size="96")
    for level in range(1, 10):
        style = ensure_style(styles_root, f"TOC{level}", f"toc {level}")
        set_font_format(ensure_child(style, RPR), "宋体", "Times New Roman")


def mark_toc_fields_dirty(container: ET.Element) -> None:
    for field in container.findall(".//w:fldChar[@w:fldCharType='begin']", NS):
        field.set(qn("dirty"), "true")


def ensure_update_fields(settings_xml: bytes | None) -> bytes | None:
    if settings_xml is None:
        return None
    settings = ET.fromstring(settings_xml)
    ensure_child(settings, qn("updateFields")).set(VAL, "true")
    return ET.tostring(settings, encoding="utf-8", xml_declaration=True)


def insert_before(body: ET.Element, reference: ET.Element | None, paragraph: ET.Element) -> None:
    if reference is None:
        sect_pr = next((child for child in list(body) if child.tag == qn("sectPr")), None)
        if sect_pr is None:
            body.append(paragraph)
        else:
            body.insert(list(body).index(sect_pr), paragraph)
        return
    body.insert(list(body).index(reference), paragraph)


def maintain_toc(
    root: ET.Element,
    body: ET.Element,
    styles_xml: bytes | None,
    settings_xml: bytes | None,
) -> tuple[bytes, bytes | None]:
    """Create or normalize a dynamic TOC and make Word refresh its cached entries."""
    styles = ET.fromstring(styles_xml) if styles_xml else ET.Element(qn("styles"))
    style_levels = style_outline_levels(styles_xml)
    containers = toc_containers(root)
    if not containers:
        toc = make_toc_content_control()
        paragraphs = direct_paragraphs(body)
        first_main_heading = next(
            (
                paragraph
                for paragraph in paragraphs
                if paragraph_heading_level(paragraph, style_levels) == 1
            ),
            None,
        )
        insert_before(body, first_main_heading, toc)
        containers = [toc]

    for container in containers:
        title = toc_title_paragraph(container)
        if title is not None:
            format_toc_title(title)
        for entry in toc_entry_paragraphs(container):
            format_toc_entry_paragraph(entry)
        mark_toc_fields_dirty(container)
    format_toc_styles(styles)
    return ET.tostring(styles, encoding="utf-8", xml_declaration=True), ensure_update_fields(settings_xml)


def find_soffice() -> Path | None:
    configured = Path("/Users/liuxs/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice")
    candidates = [
        Path(value)
        for value in (
            __import__("os").environ.get("SOFFICE_PATH", ""),
            shutil.which("soffice") or "",
        )
        if value
    ]
    candidates.append(configured)
    return next((candidate for candidate in candidates if candidate.is_file() and candidate.exists()), None)


def write_refresh_macro(profile: Path, document_path: Path, marker: Path) -> None:
    """Install a one-shot user-profile macro without altering the user's LO profile."""
    basic = profile / "basic" / "Standard"
    basic.mkdir(parents=True, exist_ok=True)
    (profile / "basic" / "script.xlc").write_text(
        """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<!DOCTYPE library:libraries PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"libraries.dtd\">
<library:libraries xmlns:library=\"http://openoffice.org/2000/library\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">
 <library:library library:name=\"Standard\" library:link=\"false\"/>
</library:libraries>
""",
        encoding="utf-8",
    )
    (basic / "script.xlb").write_text(
        """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<!DOCTYPE library:library PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"library.dtd\">
<library:library xmlns:library=\"http://openoffice.org/2000/library\" library:name=\"Standard\" library:readonly=\"false\" library:passwordprotected=\"false\">
 <library:element library:name=\"Module1\"/>
</library:library>
""",
        encoding="utf-8",
    )

    def basic_string(path: Path) -> str:
        return str(path).replace('"', '""')

    document_url = document_path.resolve().as_uri().replace('"', '""')
    marker_path = basic_string(marker)
    (basic / "Module1.xba").write_text(
        f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">Sub Main()
On Error GoTo Failed
Dim args(1) As New com.sun.star.beans.PropertyValue
args(0).Name = "Hidden"
args(0).Value = True
args(1).Name = "UpdateDocMode"
args(1).Value = 3
Dim doc As Object
doc = StarDesktop.loadComponentFromURL("{document_url}", "_blank", 0, args())
Dim i As Integer
For i = 0 To doc.DocumentIndexes.Count - 1
  doc.DocumentIndexes.getByIndex(i).update()
Next i
doc.store()
doc.close(True)
Open "{marker_path}" For Output As #1
Print #1, "updated"
Close #1
Exit Sub
Failed:
Open "{marker_path}" For Output As #1
Print #1, "failed:" & Err & ":" & Error$
Close #1
End Sub</script:module>
''',
        encoding="utf-8",
    )


def refresh_toc_with_libreoffice(document_path: Path) -> bool:
    """Return whether a local LibreOffice installation refreshed cached TOC entries."""
    soffice = find_soffice()
    if soffice is None:
        return False
    with tempfile.TemporaryDirectory(prefix="word-qa-s-toc-") as directory:
        work = Path(directory)
        refresh_copy = work / "document.docx"
        marker = work / "updated.txt"
        profile = work / "profile"
        shutil.copy2(document_path, refresh_copy)
        write_refresh_macro(profile, refresh_copy, marker)
        try:
            completed = subprocess.run(
                [
                    str(soffice),
                    "--headless",
                    "--nologo",
                    "--nodefault",
                    "--nofirststartwizard",
                    f"-env:UserInstallation={profile.resolve().as_uri()}",
                    str(refresh_copy),
                    "macro:///Standard.Module1.Main()",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=10,
                check=False,
            )
        except subprocess.TimeoutExpired:
            return False
        marker_text = marker.read_text(encoding="utf-8") if marker.is_file() else ""
        if completed.returncode != 0 or marker_text.strip() != "updated":
            return False
        if not zipfile.is_zipfile(refresh_copy):
            return False
        shutil.copy2(refresh_copy, document_path)
        return True


def process_group(
    body: ET.Element,
    scope: AnswerScope,
    content: list[ET.Element],
    headings: list[Heading],
    insert_at: ET.Element | None,
) -> list[str]:
    issues: list[str] = []
    texts = [paragraph_text(p) for p in content]
    word_texts = [word_paragraph_text(p) for p in content]
    heading_by_paragraph = {heading.paragraph: heading for heading in headings}
    local_headings = [
        (index, heading_by_paragraph[paragraph])
        for index, paragraph in enumerate(content)
        if paragraph in heading_by_paragraph
    ]
    all_labelled_question_indices = [
        index
        for index, (paragraph, text) in enumerate(zip(content, word_texts))
        if paragraph not in heading_by_paragraph and BRACKETED_QUESTION_RE.match(text)
    ]
    label_mode = len(all_labelled_question_indices) >= 2
    raw_question_indices = [
        index
        for index, (paragraph, text) in enumerate(zip(content, word_texts))
        if paragraph not in heading_by_paragraph
        and (
            BRACKETED_QUESTION_RE.match(text)
            if label_mode
            else question_start_match(text)
        )
    ]

    def heading_path(question_index: int) -> list[Heading]:
        path: list[Heading] = []
        ceiling = float("inf")
        for index, heading in reversed(local_headings):
            if index >= question_index or heading.level <= scope.title.level:
                continue
            if heading.level < ceiling:
                path.append(heading)
                ceiling = heading.level
            elif heading.level == ceiling and path and path[-1].is_type and not heading.is_type:
                # Some source documents assign a section title and its following
                # question-type title the same Word outline level. Keep the
                # nearest non-type title as the type heading's parent.
                path.append(heading)
        return list(reversed(path))

    raw_question_paths = {index: heading_path(index) for index in raw_question_indices}
    raw_question_numbers: dict[int, str] = {}
    synthetic_question_numbers: dict[int, str] = {}
    next_number_by_path: dict[tuple[ET.Element, ...], int] = {}
    for index in raw_question_indices:
        path_key = tuple(heading.paragraph for heading in raw_question_paths[index])
        explicit_number = question_number_from_text(word_texts[index])
        if explicit_number is not None:
            raw_question_numbers[index] = explicit_number
            next_number_by_path[path_key] = max(
                next_number_by_path.get(path_key, 1), int(explicit_number) + 1
            )
            continue
        generated_number = str(next_number_by_path.get(path_key, 1))
        raw_question_numbers[index] = generated_number
        synthetic_question_numbers[index] = generated_number
        next_number_by_path[path_key] = int(generated_number) + 1
    numbered_table_entries: dict[int, tuple[list[Heading], str, list[ET.Element]]] = {}
    numbered_answer_rows: set[int] = set()
    numbered_table_positions: set[int] = set()
    invalid_numbered_table_positions: set[int] = set()
    moved: list[ET.Element] = []
    source_name = scope.title.normalized or scope.title.text

    for marker_index in ([] if label_mode else range(len(content) - 1)):
        marker_match = answer_prefix_match(texts[marker_index])
        if marker_match is None or marker_match.end() != len(texts[marker_index]):
            continue
        row_indices: list[int] = []
        next_index = marker_index + 1
        while next_index < len(content) and QUESTION_RE.match(word_texts[next_index]):
            row_indices.append(next_index)
            next_index += 1
        if len(row_indices) < 2:
            continue
        answer_numbers = [QUESTION_RE.match(word_texts[index]).group(1) for index in row_indices]
        expected_numbers = [str(number) for number in range(1, len(row_indices) + 1)]
        answer_path = tuple(heading.paragraph for heading in raw_question_paths[row_indices[0]])
        preceding_questions = [
            index
            for index in raw_question_indices
            if index < marker_index
            and tuple(heading.paragraph for heading in raw_question_paths[index]) == answer_path
        ]
        matched_questions = preceding_questions[-len(row_indices) :]
        valid = (
            answer_numbers == expected_numbers
            and len(matched_questions) == len(row_indices)
            and [raw_question_numbers[index] for index in matched_questions] == expected_numbers
        )
        if not valid:
            invalid_numbered_table_positions.update({marker_index, *row_indices})
            issues.append(f"{source_name}：编号答案表无法可靠拆分，未移动。")
            continue
        for question_index, row_index, number in zip(matched_questions, row_indices, expected_numbers):
            numbered_table_entries[question_index] = (
                raw_question_paths[question_index],
                number,
                [clone_numbered_answer_row(content[row_index])],
            )
        numbered_answer_rows.update(row_indices)
        numbered_table_positions.update({marker_index, *row_indices})
        moved.extend(content[index] for index in (marker_index, *row_indices))

    question_indices = [index for index in raw_question_indices if index not in numbered_answer_rows]
    question_numbers = {index: raw_question_numbers[index] for index in question_indices}
    question_paths = {index: raw_question_paths[index] for index in question_indices}
    range_entries: dict[int, tuple[list[Heading], str, list[ET.Element]]] = {}
    range_question_indices: set[int] = set()
    range_positions: set[int] = set()
    invalid_range_positions: set[int] = set()

    for marker_index in ([] if label_mode else range(len(content) - 1)):
        marker_match = answer_prefix_match(texts[marker_index])
        if marker_match is None or marker_match.end() != len(texts[marker_index]):
            continue
        range_indexes: list[int] = []
        next_index = marker_index + 1
        while next_index < len(content) and RANGE_ANSWER_RE.fullmatch(texts[next_index]):
            range_indexes.append(next_index)
            next_index += 1
        if not range_indexes:
            continue
        staged_entries: dict[int, tuple[list[Heading], str, list[ET.Element]]] = {}
        claimed_indices: set[int] = set()
        valid = True
        preceding_questions = [index for index in question_indices if index < marker_index]
        anchor_path = (
            tuple(heading.paragraph for heading in question_paths[preceding_questions[-1]])
            if preceding_questions
            else None
        )
        for range_index in range_indexes:
            start, end, letters = RANGE_ANSWER_RE.fullmatch(texts[range_index]).groups()
            start_number = int(start)
            end_number = int(end)
            expected_numbers = [str(number) for number in range(start_number, end_number + 1)]
            matched_indices = [
                index
                for index in preceding_questions
                if question_numbers[index] in expected_numbers
                and tuple(heading.paragraph for heading in question_paths[index]) == anchor_path
            ]
            range_valid = (
                anchor_path is not None
                and start_number <= end_number
                and len(letters) == len(expected_numbers)
                and len(matched_indices) == len(expected_numbers)
                and [question_numbers[index] for index in matched_indices] == expected_numbers
                and not claimed_indices.intersection(matched_indices)
                and len({tuple(heading.paragraph for heading in question_paths[index]) for index in matched_indices}) == 1
            )
            if not range_valid:
                valid = False
                issues.append(f"{source_name}：范围答案“{texts[range_index][:60]}”无法可靠拆分，未移动。")
                continue
            claimed_indices.update(matched_indices)
            for index, letter in zip(matched_indices, letters):
                staged_entries[index] = (
                    question_paths[index],
                    question_numbers[index],
                    [compact_answer_paragraph(content[marker_index], letter)],
                )
        if not valid:
            invalid_range_positions.update({marker_index, *range_indexes})
            continue
        range_entries.update(staged_entries)
        range_question_indices.update(claimed_indices)
        range_positions.update({marker_index, *range_indexes})
        moved.extend(content[index] for index in (marker_index, *range_indexes))

    inline_entries: dict[int, tuple[list[Heading], str, list[ET.Element]]] = {}
    inline_question_positions: set[int] = set()
    for index in question_indices:
        if index in numbered_table_entries or index in range_question_indices:
            continue
        question_match = question_start_match(word_texts[index])
        marker_match = inline_answer_prefix_match(texts[index], question_match)
        if marker_match is None:
            continue
        answer_fragment = paragraph_text_fragment(
            content[index], marker_match.start(), len(texts[index])
        )
        question_fragment = paragraph_text_fragment(content[index], 0, marker_match.start())
        inline_entries[index] = (
            question_paths[index],
            question_numbers[index],
            [answer_fragment],
        )
        replace_paragraph_contents(content[index], question_fragment)
        if index in synthetic_question_numbers:
            prepend_source_question_number(content[index], synthetic_question_numbers[index])
        inline_question_positions.add(index)

    explicit_question_indices = (
        set(numbered_table_entries) | range_question_indices | set(inline_entries)
    )
    for q_pos, index in enumerate(question_indices):
        if index in explicit_question_indices:
            continue
        next_question = question_indices[q_pos + 1] if q_pos + 1 < len(question_indices) else len(content)
        next_heading = next(
            (pos for pos in range(index + 1, next_question) if content[pos] in heading_by_paragraph),
            next_question,
        )
        end = min(next_question, next_heading)
        candidates = [
            pos
            for pos in range(index + 1, end)
            if (
                pos not in invalid_range_positions
                and pos not in range_positions
                and pos not in invalid_numbered_table_positions
                and pos not in numbered_table_positions
                and is_answer_candidate(texts[pos])
            )
        ]
        if candidate_positions_are_a_single_answer_block(
            candidates,
            texts,
            allow_repeated_canonical_subanswers=label_mode,
        ):
            explicit_question_indices.add(index)
    allow_markerless_answers = (
        len(question_indices) >= 2
        and len(explicit_question_indices) >= 2
        and len(explicit_question_indices) / len(question_indices) >= 0.8
    )

    entries: list[tuple[list[Heading], str, list[ET.Element]]] = []
    split_answer_source_positions: set[int] = set()

    for q_pos, index in enumerate(question_indices):
        if index in numbered_table_entries:
            entries.append(numbered_table_entries[index])
            continue
        if index in range_question_indices:
            entries.append(range_entries[index])
            continue
        if index in inline_entries:
            entries.append(inline_entries[index])
            continue
        number = question_numbers[index]
        next_question = question_indices[q_pos + 1] if q_pos + 1 < len(question_indices) else len(content)
        next_heading = next(
            (pos for pos in range(index + 1, next_question) if content[pos] in heading_by_paragraph),
            next_question,
        )
        end = min(next_question, next_heading)
        candidates = [
            pos
            for pos in range(index + 1, end)
            if (
                pos not in invalid_range_positions
                and pos not in range_positions
                and pos not in invalid_numbered_table_positions
                and pos not in numbered_table_positions
                and is_answer_candidate(texts[pos])
            )
        ]
        if candidate_positions_are_a_single_answer_block(
            candidates,
            texts,
            allow_repeated_canonical_subanswers=label_mode,
        ):
            answer_start = candidates[0]
            answer_end = next(
                (pos for pos in range(answer_start + 1, end) if content[pos] in heading_by_paragraph),
                end,
            )
            embedded_marker = (
                None
                if answer_prefix_match(texts[answer_start]) is not None
                else INLINE_ANSWER_PREFIX_RE.search(texts[answer_start])
            )
            if embedded_marker is not None:
                answer_head = paragraph_text_fragment(
                    content[answer_start], embedded_marker.start(), len(texts[answer_start])
                )
                retained_head = paragraph_text_fragment(
                    content[answer_start], 0, embedded_marker.start()
                )
                replace_paragraph_contents(content[answer_start], retained_head)
                answer_block = [answer_head, *content[answer_start + 1 : answer_end]]
                split_answer_source_positions.add(answer_start)
                moved.extend(content[answer_start + 1 : answer_end])
            else:
                answer_block = content[answer_start:answer_end]
                moved.extend(answer_block)
            if index in synthetic_question_numbers:
                prepend_source_question_number(content[index], synthetic_question_numbers[index])
            entries.append((question_paths[index], number, answer_block))
        elif len(candidates) == 0:
            answer_start = markerless_answer_start(texts, index, end)
            if allow_markerless_answers and answer_start is not None:
                if index in synthetic_question_numbers:
                    prepend_source_question_number(content[index], synthetic_question_numbers[index])
                entries.append(
                    (
                        question_paths[index],
                        number,
                        clone_markerless_answer_block(content[answer_start:end]),
                    )
                )
                moved.extend(content[index + 1 : end])
            else:
                question_preview = re.sub(r"\s+", " ", texts[index]).strip()[:40]
                issues.append(
                    f"{source_name}：第 {number} 题未找到紧随的 {ANSWER_MARKER}"
                    f"（题干“{question_preview}”）。"
                )
        else:
            issues.append(f"{source_name}：第 {number} 题找到 {len(candidates)} 条 {ANSWER_MARKER}，未移动。")

    assigned = set(moved) | {
        content[index] for index in inline_question_positions | split_answer_source_positions
    }
    for paragraph, text in zip(content, texts):
        if is_answer_candidate(text) and paragraph not in assigned:
            # A pairing can be ambiguous, but an explicit leading alias is not:
            # normalize it in place so no visible `答：/解：/解析：` prefix is
            # left behind.  The review record still prevents unsafe movement.
            prefix_match = answer_prefix_match(text)
            if prefix_match is not None and not text.lstrip().startswith(ANSWER_MARKER):
                replace_leading_answer_prefix(paragraph)
            issues.append(f"{source_name}：未能可靠配对的答案“{text[:60]}”。")

    if not entries:
        return issues
    for paragraph in moved:
        body.remove(paragraph)

    answer_title = clone_text_paragraph(scope.title.paragraph, scope.answer_title)
    insert_before(body, insert_at, answer_title)
    last_path: list[Heading] = []
    for path, number, answer_block in entries:
        common_length = 0
        while (
            common_length < len(last_path)
            and common_length < len(path)
            and last_path[common_length].paragraph is path[common_length].paragraph
        ):
            common_length += 1
        for heading in path[common_length:]:
            copied_heading = (
                clone_heading_with_advanced_number(heading.paragraph)
                if heading.is_type
                else copy.deepcopy(heading.paragraph)
            )
            insert_before(body, insert_at, copied_heading)
        last_path = path
        for answer_paragraph in clone_answer_block(answer_block, number):
            insert_before(body, insert_at, answer_paragraph)
    return issues


def separate_answers(
    input_path: Path,
    output_path: Path,
    force: bool,
    *,
    refresh_toc: bool = False,
) -> tuple[int, list[str]]:
    if input_path.suffix.lower() != ".docx":
        raise ValueError("仅支持 .docx 文件。")
    if not input_path.is_file():
        raise FileNotFoundError(f"找不到输入文件：{input_path}")
    if output_path.exists() and not force:
        raise FileExistsError(f"输出文件已存在：{output_path}（如需覆盖，请使用 --force）")

    with zipfile.ZipFile(input_path) as source:
        try:
            document_xml = source.read("word/document.xml")
        except KeyError as exc:
            raise ValueError("该文件不是包含 word/document.xml 的有效 Word 文档。") from exc
        root = ET.fromstring(document_xml)
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("文档缺少正文。")

        styles_xml = source.read("word/styles.xml") if "word/styles.xml" in source.namelist() else None
        settings_xml = source.read("word/settings.xml") if "word/settings.xml" in source.namelist() else None
        remove_legacy_font_warning_review(body)
        split_joined_answer_and_question_paragraphs(body)
        paragraphs = direct_paragraphs(body)
        headings = collect_headings(paragraphs, style_outline_levels(styles_xml))
        restore_omitted_question_numbers(body, headings)
        paragraphs = direct_paragraphs(body)
        headings = collect_headings(paragraphs, style_outline_levels(styles_xml))
        scopes = collect_answer_scopes(headings, paragraphs)
        issues: list[str] = []
        for scope in scopes:
            current = direct_paragraphs(body)
            title_index = current.index(scope.title.paragraph)
            end_index = scope_end(scope, headings, paragraphs)
            insert_at = paragraphs[end_index] if end_index < len(paragraphs) else None
            if insert_at is not None:
                following_title = paragraph_text(insert_at)
                if following_title == scope.answer_title or is_answer_heading(following_title):
                    continue
            issues.extend(
                process_group(
                    body,
                    scope,
                    current[title_index + 1 : current.index(insert_at) if insert_at is not None else len(current)],
                    headings,
                    insert_at,
                )
            )

        if not scopes:
            issues.append("未找到带题目语义的标题，未移动任何答案。")
        new_review_issues: list[str] = []
        if issues:
            existing_paragraphs = direct_paragraphs(body)
            existing_texts = {paragraph_text(paragraph) for paragraph in existing_paragraphs}
            new_review_issues = [issue for issue in issues if issue not in existing_texts]
            if new_review_issues:
                review_template = scopes[0].title.paragraph if scopes else (existing_paragraphs[0] if existing_paragraphs else ET.Element(P))
                if "需人工复核" not in existing_texts:
                    insert_before(body, None, clone_text_paragraph(review_template, "需人工复核"))
                for issue in new_review_issues:
                    insert_before(body, None, clone_text_paragraph(review_template, issue))

        for paragraph in direct_paragraphs(body):
            normalize_answer_marker_fonts(paragraph)

        styles_xml, settings_xml = maintain_toc(
            root,
            body,
            styles_xml,
            settings_xml,
        )
        updated_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        replacements = {
            "word/document.xml": updated_xml,
            "word/styles.xml": styles_xml,
        }
        if settings_xml is not None:
            replacements["word/settings.xml"] = settings_xml
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output_path.parent, suffix=".docx", delete=False) as temporary:
            temporary_path = Path(temporary.name)
        try:
            with zipfile.ZipFile(temporary_path, "w", zipfile.ZIP_DEFLATED) as destination:
                written: set[str] = set()
                for info in source.infolist():
                    destination.writestr(info, replacements.get(info.filename, source.read(info.filename)))
                    written.add(info.filename)
                for name, data in replacements.items():
                    if name not in written:
                        destination.writestr(name, data)
            if refresh_toc:
                refresh_toc_with_libreoffice(temporary_path)
            temporary_path.replace(output_path)
        finally:
            temporary_path.unlink(missing_ok=True)
    return len(new_review_issues), issues


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}-题目答案分离.docx")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="原始 .docx 题库文件")
    parser.add_argument("--output", type=Path, help="输出 .docx 路径")
    parser.add_argument("--force", action="store_true", help="覆盖已有输出文件")
    args = parser.parse_args()
    output = args.output or default_output_path(args.input)
    try:
        issue_count, _ = separate_answers(args.input, output, args.force, refresh_toc=True)
    except (ValueError, FileNotFoundError, FileExistsError, zipfile.BadZipFile) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    print(f"已生成：{output}")
    print(f"人工复核项：{issue_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
