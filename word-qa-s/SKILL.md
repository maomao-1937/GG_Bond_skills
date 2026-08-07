 ---
name: word-qa-s
description: "Separate questions from their immediately following answers in Word (.docx) question-bank documents. Use when a user asks to move `【万人教育解析】...` answer paragraphs out of question text and create per-chapter or per-section answer areas, including question groups with numbered or varied headings."
---

# Word 题库答案分离

Use this skill only for `.docx` question-bank documents. Do not process legacy `.doc`, images, PDFs, or pasted plain text as if they were Word files.

## Workflow

1. Confirm the input is a `.docx` file. Keep it unchanged.
2. Run the bundled script. It writes a sibling copy whose name ends in `-题目答案分离.docx`.

   ```bash
   python3 scripts/separate_answers.py "/path/to/题库.docx"
   ```

   Pass `--output "/path/to/结果.docx"` to choose the destination. Refuse to overwrite an existing file unless the user explicitly requests `--force`.

3. Inspect the generated document. Confirm each recognized question group retains its questions, has a following answer area named after the source title, has no moved answer left among the questions, and has an up-to-date table of contents.
4. Tell the user the output path and call out the `需人工复核` section if present.

## Processing Rules

- Recognize question-group headings by Word outline level plus question wording. Numbered forms such as `4.本章例题` are valid, as are headings containing `例题、习题、练习题、题目、测试题、自测题、思考题` and common question types. If a chapter/section title has no question keyword, infer a chapter-level answer area only from a substantial consistent sample: at least two numbered questions, each with exactly one recognized answer marker before the next numbered question or heading, and at least 80% of the chapter's numbered questions satisfying that pairing. This supports ordinary titles such as `第二章 8086 CPU 的结构与功能` without treating ordinary prose as a question bank.
- If an inferred chapter does not meet that threshold, evaluate its directly delimited sections independently. Create a section answer area only for a qualifying section, and preserve its original numeric title (for example, `12.1…` remains `12.1…（题目答案）` rather than being advanced to the next chapter number).
- Read heading levels from the Word style definition first (for example, `heading 1/2/3`); use a numeric-title fallback only when style hierarchy is unavailable.
- Merge nested question-type headings such as `选择题` and `简答题` into their parent question group. Do not create a separate answer area for each type.
- Derive the answer-area title from the source title and keep its style: replace `（题目）` with `（题目答案）`; append `答案` to headings ending in `例题` or `习题`; otherwise append `（题目答案）`. For example, `第二讲 三种调速方法简介（题目）` becomes `第二讲 三种调速方法简介（题目答案）`.
- When a source question heading begins with an Arabic dotted number, advance its first number only in the generated answer area. For example, `4.本章例题` / `4.1选择题` become `5.本章例题答案` / `5.1选择题`; unnumbered headings are unchanged.
- Do not treat an existing generated answer heading (including legacy `本章/本节习题答案`) as source content; rerunning the script must not duplicate answer areas.
- Identify questions that start with `1.` / `1．` / `1、` (and the equivalent for any number), and the explicit teaching labels `【例】` and `【题n】`. Treat `6.8086 CPU…` as question 6 when the digits after the dot introduce a title/acronym, but do not treat a decimal calculation such as `11010.1=26.5` as a question. `【题n】` retains `n`; a marker-bounded unnumbered `【例】` receives a manual local number only after its answer pairing is proven. Labels such as `【注】`、`【做题步骤】`、`【理想点评】` are never questions.
- Read both normal Word text and Office Math text when detecting headings, question numbers, and answer markers. Treat a paragraph beginning with `【万人教育解析】`, `答：`, `答案：`, `解：`, `解答：`, `解析：`, `分析：`, `题目分析：`, or a subanswer label such as `题（1）解：` (Chinese or ASCII colon) as an answer marker, including when the marker is rendered as an equation. Normalize every moved answer's first paragraph to the exact form `题号.【万人教育解析】答案内容`, with no space after the marker. Strip redundant answer aliases/markers from continuation paragraphs in the same answer block so only its first paragraph carries the canonical numbered label.
- Before pairing, split a marked paragraph that embeds one or more subsequent question numbers or a question-type title (for example `…【万人教育解析】B。2、下一题` or `答：A。2、下一题`). Support the canonical marker and all accepted aliases during this split. Partition the OOXML children instead of copying the whole paragraph: every drawing, formula, OLE object, and other non-text payload belongs to exactly one resulting fragment. Insert the fragments at their true XML sibling location so tables and objects cannot jump or duplicate.
- Also split an explicit `【例】` or `【题n】` label joined to the end of a preceding sentence (for example `方法如下。【题1】求解…`), but only after sentence-ending punctuation; this makes the label a true question boundary without treating ordinary bracketed notes as questions.
- Restore an omitted source question number only after Word heading levels have been resolved. Every Word heading is a numbering boundary: a first recoverable question under a new heading starts at `1`, and only another question in the same heading/type scope may inherit the previous number plus one. Require a following explicit answer marker before the next numbered question or heading. Never promote a bare assembly-code underline or step placeholder to a question; when the number cannot be proved, preserve the source and add a review item.
- Split a same-paragraph `题干【万人教育解析】答案` form into the retained question and a canonical answer entry. Do not mistake ordinary wording such as `请回答：` for an inline `答：` marker. If an alias prefix cannot be paired safely, replace only that leading alias with `【万人教育解析】` in place and add a `需人工复核` item; do not guess a question number or move the content.
- Move an individual answer when its recognized marker block is unambiguous before the next numbered question. A canonical marker followed only by alias-style explanation paragraphs is one answer block. Move the complete answer block through the paragraph before the next question number or heading boundary, including explanatory paragraphs, equations, images, embedded objects, and original formatting. Internally identify questions by their paragraph occurrence/XML node; the visible number is only output text. Two questions displaying the same number are therefore never allowed to cancel each other's missing-answer record or borrow one answer.
- In a `【例】/【题n】` label-based scope, numbered lines inside a marked solution are treated as solution steps, not new questions. Repeated canonical markers are one answer block only when each later marker is introduced by an explicit subanswer label such as `（2）`; otherwise retain the ambiguity for review.
- For an inferred chapter scope, retain the intermediate section/subsection headings in the answer area and use one chapter-wide area rather than overlapping answer areas for nested unlabelled sections. Do not infer a scope if the question sits inside an existing generated answer area or if there is no marker-bounded numbered question.
- Expand a compact range answer such as `1-5：BAAAD` only when the referenced question numbers are consecutive, all occur under the same heading chain, and the answer string contains exactly one character per question. Generate `1.【万人教育解析】B` through `5.【万人教育解析】D`; do not infer multi-select or count-mismatched answers.
- Split a numbered answer table only when it immediately follows an answer marker and contains a complete consecutive sequence such as `1、答案甲` through `15、答案乙`. Match it to the immediately preceding question sequence under the same heading chain, remove the marker and source answer rows from the question area, and generate one canonical answer per question. Do not treat an unmarked numbered list as answers.
- Preserve the complete heading chain leading to each moved answer in the generated answer area. Copy section headings (for example, `1.2、1.3、1.4微机中的数制与码制`) verbatim and preserve question-type headings, including numbered forms such as `一、选择题`.
- Preserve the original formatting of every answer run. The exact label `【万人教育解析】` is the exception: place it in its own Word run where possible and set its Chinese/Latin font attributes to 华文新魏. The prepended `题号.` remains in a separate run using the ordinary source/body font, never 华文新魏.
- An unmarked explanatory block may be moved only inside an already established question scope where at least two questions have explicit markers and at least 80% of the scope's questions are explicitly paired. Unmarked prose never counts toward establishing the scope itself. Always exclude reference/material content beginning with `规格、控制信号、格式如下、主要包括`.
- Leave genuinely ambiguous, count-mismatched, nonconsecutive, duplicated-number, or unmatched content in place and add a truthful `需人工复核` item. Never suppress a missing record merely because another question under the same headings displays the same number, and never force the review count to zero.
- Treat a Word dynamic TOC content control or `TOC` field as the document's directory. Keep it in place, update its field state, and normalize its title and entries after every answer separation.
- If no directory exists, add a dynamic TOC immediately before the first level-one body heading and put a page break after it. The field covers levels 1–3 and is marked for Word to update on open.
- Format the directory title `目录` as 华文新魏, centered, 48 pt, and not bold. Format every TOC entry with Chinese in 宋体, Latin letters/numbers in Times New Roman, and no bold formatting.
- Do not add a review page merely because a requested font is not installed locally. `需人工复核` is reserved for genuine answer-pairing ambiguity.

## Verification

After generating a document, validate its package and render it when the result will be delivered to a user:

```bash
python3 /Users/liuxs/.agents/skills/docx/scripts/office/validate.py output.docx --original input.docx
python3 /Users/liuxs/.agents/skills/docx/scripts/office/soffice.py --headless --convert-to pdf output.docx
```
