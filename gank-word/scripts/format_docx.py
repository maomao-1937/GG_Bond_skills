#!/usr/bin/env python3
"""
万人教育 Word 文档格式纠正脚本

用法: python format_docx.py input.docx [-o output.docx]

功能:
1. 修改 Word 内建样式定义（Heading1/2/3, Normal, Header, Footer, TOCHeading）
2. 按规范重写内建样式；缺失的样式用 Word 内建名称新建
3. 扫描全文标题层级并做全局偏移纠正
4. 按内容特征识别"解析"段落：标签 run 保持华文新魏，正文继承 Normal（宋体）
5. 清除所有段落的自动编号 (w:numPr)
6. 设置页眉/页脚距边界距离
"""

import argparse
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

# Word Open XML namespace
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

# Register namespaces so ET doesn't mangle prefixes
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

# Also register common namespaces found in docx that we don't actively use
_EXTRA_NS = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}
for prefix, uri in _EXTRA_NS.items():
    ET.register_namespace(prefix, uri)


# ─── Style definitions ────────────────────────────────────────────────────────

STYLE_DEFS = {
    'TOCHeading': {
        'pPr': {'jc': 'center'},
        'rPr': {
            'rFonts_eastAsia': '华文新魏',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '96', 'szCs': '96',
            'color': '000000',
            'bold': False,
        },
    },
    'Heading1': {
        'pPr': {'jc': 'center', 'spacing_before': '350', 'spacing_after': '320'},
        'rPr': {
            'rFonts_eastAsia': '宋体',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '32', 'szCs': '32',
            'color': '000000',
            'bold': True,
        },
    },
    'Heading2': {
        'pPr': {'jc': 'left', 'spacing_line': '360', 'spacing_lineRule': 'auto',
                'spacing_before': '0', 'spacing_after': '0'},
        'rPr': {
            'rFonts_eastAsia': '宋体',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '30', 'szCs': '30',
            'color': '000000',
            'bold': True,
        },
    },
    'Heading3': {
        'pPr': {'jc': 'left', 'spacing_line': '360', 'spacing_lineRule': 'auto'},
        'rPr': {
            'rFonts_eastAsia': '宋体',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '28', 'szCs': '28',
            'color': '000000',
            'bold': True,
        },
    },
    'Normal': {
        'pPr': {'jc': 'both', 'spacing_line': '360', 'spacing_lineRule': 'auto'},
        'rPr': {
            'rFonts_eastAsia': '宋体',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '24', 'szCs': '24',
            'color': '000000',
            'bold': False,
        },
    },
    'Header': {
        'pPr': {'jc': 'center'},
        'rPr': {
            'rFonts_eastAsia': '华文新魏',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '21', 'szCs': '21',
            'color': '000000',
            'bold': False,
        },
    },
    'Footer': {
        'pPr': {'jc': 'center'},
        'rPr': {
            'rFonts_eastAsia': '华文新魏',
            'rFonts_ascii': 'Times New Roman',
            'rFonts_hAnsi': 'Times New Roman',
            'sz': '21', 'szCs': '21',
            'color': '000000',
            'bold': False,
        },
    },
    # NOTE: 解析 paragraphs are NOT a dedicated paragraph style. They use
    # Normal (宋体); only the run(s) covering the 【万人教育解析】 label keep
    # 华文新魏. See mark_label_runs() / Phase 3.
}

# Font specification for TOC entry styles (toc 1..toc 9). The entries use the
# same body font (宋体 / Times New Roman) — NOT the decorative 华文新魏 that the
# "目录" heading title uses.  The style update is a *targeted* rFonts change:
# spacing, indentation, tab stops, size and color are left intact so the
# generated TOC layout (indents, dot leaders, page numbers) is preserved.
TOC_ENTRY_FONT = {
    'rFonts_eastAsia': '宋体',
    'rFonts_ascii': 'Times New Roman',
    'rFonts_hAnsi': 'Times New Roman',
}

# Known heading style IDs (case-insensitive matching)
HEADING_STYLE_PATTERNS = {
    1: re.compile(r'^(heading\s*1|标题\s*1|1)$', re.IGNORECASE),
    2: re.compile(r'^(heading\s*2|标题\s*2|2)$', re.IGNORECASE),
    3: re.compile(r'^(heading\s*3|标题\s*3|3)$', re.IGNORECASE),
    4: re.compile(r'^(heading\s*4|标题\s*4|4)$', re.IGNORECASE),
    5: re.compile(r'^(heading\s*5|标题\s*5|5)$', re.IGNORECASE),
    6: re.compile(r'^(heading\s*6|标题\s*6|6)$', re.IGNORECASE),
}

TOC_HEADING_PATTERN = re.compile(r'^(tocheading|toc\s*heading|目录标题)$', re.IGNORECASE)

# Auto-generated TOC entry styles (toc 1 .. toc 9). These paragraphs are field
# output; Word regenerates them. We only rewrite their *font* (宋体) and strip
# bold — indents and page-number tab leaders are preserved untouched.
TOC_ENTRY_PATTERN = re.compile(r'^(toc\s*[1-9]|目录\s*[1-9])$', re.IGNORECASE)

# Name patterns for the non-heading built-ins, so we can resolve a document's
# real style IDs (e.g. Normal is often styleId 'a' in CN-authored files).
BUILTIN_NAME_PATTERNS = {
    'Normal': re.compile(r'^(normal|正文)$', re.IGNORECASE),
    'Header': re.compile(r'^(header|页眉)$', re.IGNORECASE),
    'Footer': re.compile(r'^(footer|页脚)$', re.IGNORECASE),
}

# Localized built-in names to use when a style is created from scratch, so Word
# recognizes it as a built-in (TOC regeneration, navigation, UI labels). Names
# only matter on creation; existing styles keep whatever name they have.
BUILTIN_NAME_MAP = {
    'TOCHeading': '目录标题',
    'Heading1': '标题 1',
    'Heading2': '标题 2',
    'Heading3': '标题 3',
    'Normal': '正文',
    'Header': '页眉',
    'Footer': '页脚',
}

HEADING_LOGICAL_IDS = {1: 'Heading1', 2: 'Heading2', 3: 'Heading3'}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def w(tag):
    """Create a namespaced tag name."""
    return f'{{{NS["w"]}}}{tag}'


# ─── Namespace-safe (de)serialization ─────────────────────────────────────────
#
# ElementTree only emits xmlns declarations for namespaces it actually sees in
# use. Word parts declare ~35 prefixes on the root and then list a subset in
# mc:Ignorable. If a prefix named in mc:Ignorable is not declared, the file is
# not namespace-well-formed and Word reports it as unreadable. So we snapshot
# the original root start tag and restore any declaration ET dropped.

XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

_XMLNS_RE = re.compile(r'xmlns:([A-Za-z_][\w.\-]*)\s*=\s*"([^"]*)"')


def _find_root_start_tag(data):
    """Return (start, end, text) of the first element start tag in `data`."""
    i = 0
    n = len(data)
    while i < n:
        lt = data.find(b'<', i)
        if lt == -1:
            return None
        nxt = data[lt + 1:lt + 2]
        if nxt == b'?':
            end = data.find(b'?>', lt)
            i = (end + 2) if end != -1 else n
            continue
        if nxt == b'!':
            if data.startswith(b'<!--', lt):
                end = data.find(b'-->', lt)
                i = (end + 3) if end != -1 else n
            else:
                end = data.find(b'>', lt)
                i = (end + 1) if end != -1 else n
            continue
        # A real element start tag. Scan to its '>' honouring quotes.
        j = lt + 1
        quote = None
        while j < n:
            c = data[j:j + 1]
            if quote:
                if c == quote:
                    quote = None
            elif c in (b'"', b"'"):
                quote = c
            elif c == b'>':
                return lt, j + 1, data[lt:j + 1].decode('utf-8', 'replace')
            j += 1
        return None
    return None


def snapshot_namespaces(path):
    """Record the prefix→URI declarations on a part's root element.

    Also registers them globally so ET reuses the document's own prefixes
    instead of inventing ns0/ns1 names.
    """
    with open(path, 'rb') as fh:
        head = fh.read(65536)
    found = _find_root_start_tag(head)
    if not found:
        return {}
    decls = dict(_XMLNS_RE.findall(found[2]))
    for prefix, uri in decls.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    return decls


def write_tree(tree, path, ns_decls):
    """Serialize `tree`, restoring any xmlns declaration ET left out."""
    tree.write(path, xml_declaration=False, encoding='UTF-8')
    with open(path, 'rb') as fh:
        data = fh.read()

    found = _find_root_start_tag(data)
    if found and ns_decls:
        start, end, tag_text = found
        present = {p for p, _ in _XMLNS_RE.findall(tag_text)}
        missing = [(p, u) for p, u in ns_decls.items() if p not in present]
        if missing:
            extra = ''.join(f' xmlns:{p}="{u}"' for p, u in sorted(missing))
            self_closing = tag_text.endswith('/>')
            cut = -2 if self_closing else -1
            new_tag = tag_text[:cut] + extra + ('/>' if self_closing else '>')
            data = data[:start] + new_tag.encode('utf-8') + data[end:]

    with open(path, 'wb') as fh:
        fh.write(XML_DECL + data)


def find_or_create(parent, tag_name):
    """Find a child element or create it if missing."""
    el = parent.find(w(tag_name))
    if el is None:
        el = ET.SubElement(parent, w(tag_name))
    return el


def remove_child(parent, tag_name):
    """Remove a child element if present."""
    el = parent.find(w(tag_name))
    if el is not None:
        parent.remove(el)


def _clear_attrs(el):
    """Drop all attributes so stale values can't survive a rewrite.

    A style's existing <w:spacing w:before="...">, <w:rFonts w:hint="..."> or
    <w:color w:themeColor="..."> would otherwise persist alongside the values
    we set, leaving an inconsistent "unified" format.
    """
    for k in list(el.attrib):
        del el.attrib[k]


def set_rpr(rpr_el, rpr_def):
    """Apply rPr definition to an rPr element, clearing conflicting attrs."""
    # rFonts
    rfont = find_or_create(rpr_el, 'rFonts')
    _clear_attrs(rfont)
    rfont.set(w('eastAsia'), rpr_def['rFonts_eastAsia'])
    rfont.set(w('ascii'), rpr_def['rFonts_ascii'])
    rfont.set(w('hAnsi'), rpr_def['rFonts_hAnsi'])

    # sz / szCs
    sz_el = find_or_create(rpr_el, 'sz')
    _clear_attrs(sz_el)
    sz_el.set(w('val'), rpr_def['sz'])
    szcs_el = find_or_create(rpr_el, 'szCs')
    _clear_attrs(szcs_el)
    szcs_el.set(w('val'), rpr_def['szCs'])

    # color
    color_el = find_or_create(rpr_el, 'color')
    _clear_attrs(color_el)
    color_el.set(w('val'), rpr_def['color'])

    # bold
    if rpr_def['bold']:
        find_or_create(rpr_el, 'b')
        find_or_create(rpr_el, 'bCs')
    else:
        remove_child(rpr_el, 'b')
        remove_child(rpr_el, 'bCs')


def set_ppr(ppr_el, ppr_def):
    """Apply pPr definition to a pPr element."""
    # jc (alignment)
    if 'jc' in ppr_def:
        jc_el = find_or_create(ppr_el, 'jc')
        _clear_attrs(jc_el)
        jc_el.set(w('val'), ppr_def['jc'])

    # spacing: clear first so spacing attributes absent from the definition
    # (e.g. a def that only sets line/lineRule) don't leave stale before/after.
    has_spacing = any(k.startswith('spacing_') for k in ppr_def)
    if has_spacing:
        sp_el = find_or_create(ppr_el, 'spacing')
        _clear_attrs(sp_el)
        if 'spacing_before' in ppr_def:
            sp_el.set(w('before'), ppr_def['spacing_before'])
        if 'spacing_after' in ppr_def:
            sp_el.set(w('after'), ppr_def['spacing_after'])
        if 'spacing_line' in ppr_def:
            sp_el.set(w('line'), ppr_def['spacing_line'])
        if 'spacing_lineRule' in ppr_def:
            sp_el.set(w('lineRule'), ppr_def['spacing_lineRule'])


def get_style_level(style_id, style_name_map):
    """Determine heading level from a style ID. Returns None if not a heading."""
    if not style_id:
        return None
    # Check name from style_name_map first
    name = style_name_map.get(style_id, style_id)
    for level, pattern in HEADING_STYLE_PATTERNS.items():
        if pattern.match(name) or pattern.match(style_id):
            return level
    return None


def get_paragraph_text(para_el):
    """Extract plain text from a paragraph element."""
    texts = []
    for t_el in para_el.iter(w('t')):
        if t_el.text:
            texts.append(t_el.text)
    return ''.join(texts)


def build_style_name_map(styles_root):
    """Build styleId → name mapping from styles.xml."""
    name_map = {}
    for style_el in styles_root.iter(w('style')):
        sid = style_el.get(w('styleId'))
        name_el = style_el.find(w('name'))
        if sid and name_el is not None:
            name_map[sid] = name_el.get(w('val'), '')
    return name_map


def resolve_style_ids(styles_root):
    """Map our logical style keys → the styleIds this document actually uses.

    Chinese-authored documents routinely name Normal 'a', 目录标题 'af9',
    and headings '1'/'2'/'3'. Writing a literal pStyle val of 'Heading1' into
    such a file points at a style that does not exist, and Word silently
    renders the paragraph as body text. So resolve by style name first and
    only fall back to the canonical ID.
    """
    resolved = {}
    paragraph_styles = []
    for style_el in styles_root.iter(w('style')):
        if style_el.get(w('type')) not in (None, 'paragraph'):
            continue
        sid = style_el.get(w('styleId'))
        if not sid:
            continue
        name_el = style_el.find(w('name'))
        paragraph_styles.append((sid, name_el.get(w('val'), '') if name_el is not None else ''))

    by_id = {sid for sid, _ in paragraph_styles}

    # Headings: match on name, then on the bare-numeral / canonical ID.
    for level in (1, 2, 3):
        pattern = HEADING_STYLE_PATTERNS[level]
        hit = next((sid for sid, name in paragraph_styles if pattern.match(name)), None)
        if hit is None:
            hit = next((sid for sid, _ in paragraph_styles if pattern.match(sid)), None)
        resolved[HEADING_LOGICAL_IDS[level]] = hit or HEADING_LOGICAL_IDS[level]

    # TOC heading (目录标题).
    hit = next((sid for sid, name in paragraph_styles if TOC_HEADING_PATTERN.match(name)), None)
    resolved['TOCHeading'] = hit or 'TOCHeading'

    # Normal / Header / Footer.
    for key, pattern in BUILTIN_NAME_PATTERNS.items():
        hit = next((sid for sid, name in paragraph_styles if pattern.match(name)), None)
        if hit is None and key in by_id:
            hit = key
        resolved[key] = hit or key

    return resolved


def collect_toc_entry_ids(styles_root):
    """styleIds of auto-generated TOC entry styles (toc 1..toc 9)."""
    out = set()
    for style_el in styles_root.iter(w('style')):
        sid = style_el.get(w('styleId'))
        if not sid:
            continue
        name_el = style_el.find(w('name'))
        name = name_el.get(w('val'), '') if name_el is not None else ''
        if TOC_ENTRY_PATTERN.match(name) or TOC_ENTRY_PATTERN.match(sid):
            out.add(sid)
    return out


# pPr children that must follow w:outlineLvl per the OOXML schema. Appending
# outlineLvl after w:rPr (paragraph mark formatting) yields an out-of-order
# document that strict consumers / WPS may reject.
_AFTER_OUTLINE_LVL_TAGS = {w(t) for t in ('divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')}


def set_outline_lvl(ppr, val):
    """Set w:outlineLvl at its schema-correct position in pPr (before rPr)."""
    old = ppr.find(w('outlineLvl'))
    if old is not None:
        ppr.remove(old)
    el = ET.Element(w('outlineLvl'))
    el.set(w('val'), str(val))
    insert_at = next((i for i, child in enumerate(ppr)
                      if child.tag in _AFTER_OUTLINE_LVL_TAGS), None)
    if insert_at is not None:
        ppr.insert(insert_at, el)
    else:
        ppr.append(el)


def mark_label_runs(para, prefix):
    """Which runs overlap `prefix`, by text position.

    Word splits text across runs on edit, so the 【万人教育解析】 label may span
    several runs. Match on cumulative text positions rather than assuming the
    label lives in a single run.
    """
    runs = list(para.iter(w('r')))
    texts = [''.join(tx.text or '' for tx in r.iter(w('t'))) for r in runs]
    full = ''.join(texts)
    start = full.find(prefix)
    if start < 0:
        return [False] * len(runs)
    end = start + len(prefix)
    out, acc = [], 0
    for t in texts:
        r_start, r_end = acc, acc + len(t)
        acc = r_end
        out.append(r_start < end and r_end > start)
    return out


# ─── Core processing ──────────────────────────────────────────────────────────

def update_toc_entry_styles(styles_root, toc_entry_ids):
    """Set 宋体 / Times New Roman on TOC entry styles (toc 1..toc 9).

    Targeted update: only w:rFonts eastAsia/ascii/hAnsi are set, theme-font
    references that could override them are cleared, and bold (w:b / w:bCs)
    is removed.  Spacing, indentation, tab stops, size and color are left
    untouched so the generated TOC layout (indents, dot leaders, page
    numbers) survives exactly as Word authored it.
    """
    for style_el in styles_root.iter(w('style')):
        sid = style_el.get(w('styleId'))
        if not sid or sid not in toc_entry_ids:
            continue

        rpr_el = style_el.find(w('rPr'))
        if rpr_el is None:
            rpr_el = ET.SubElement(style_el, w('rPr'))

        # rFonts: targeted font-face update (do NOT wipe other attrs).
        rfont = find_or_create(rpr_el, 'rFonts')
        rfont.set(w('eastAsia'), TOC_ENTRY_FONT['rFonts_eastAsia'])
        rfont.set(w('ascii'), TOC_ENTRY_FONT['rFonts_ascii'])
        rfont.set(w('hAnsi'), TOC_ENTRY_FONT['rFonts_hAnsi'])
        # Theme-font references can take precedence over explicit faces in
        # some Word implementations — drop them so 宋体 wins.
        for theme_attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme'):
            key = w(theme_attr)
            if key in rfont.attrib:
                del rfont.attrib[key]

        # TOC entries are never bold.
        remove_child(rpr_el, 'b')
        remove_child(rpr_el, 'bCs')


def update_styles_xml(styles_path, ns_decls):
    """Rewrite style definitions in styles.xml according to STYLE_DEFS.

    Returns (resolved_ids, toc_entry_ids) so the body pass can reference the
    same styleIds this function edited.
    """
    tree = ET.parse(styles_path)
    root = tree.getroot()

    resolved = resolve_style_ids(root)
    toc_entry_ids = collect_toc_entry_ids(root)

    existing = {}
    for style_el in root.iter(w('style')):
        sid = style_el.get(w('styleId'))
        if sid:
            existing[sid] = style_el

    for logical_id, sdef in STYLE_DEFS.items():
        target_id = resolved.get(logical_id, logical_id)
        style_el = existing.get(target_id)

        if style_el is None:
            # Create a missing style with Word's built-in (localized) name so
            # Word recognizes it as a built-in for TOC/navigation/UI.
            style_el = ET.SubElement(root, w('style'))
            style_el.set(w('type'), 'paragraph')
            style_el.set(w('styleId'), target_id)
            name_sub = ET.SubElement(style_el, w('name'))
            name_sub.set(w('val'), BUILTIN_NAME_MAP.get(logical_id, logical_id))
            existing[target_id] = style_el

        ppr_el = find_or_create(style_el, 'pPr')
        set_ppr(ppr_el, sdef['pPr'])

        rpr_el = find_or_create(style_el, 'rPr')
        set_rpr(rpr_el, sdef['rPr'])

    # TOC entry styles (toc 1..toc 9): set font to 宋体 / Times New Roman
    # and remove bold; everything else about them is left alone.
    update_toc_entry_styles(root, toc_entry_ids)

    write_tree(tree, styles_path, ns_decls)
    return resolved, toc_entry_ids


def process_document_xml(doc_path, styles_path, resolved, toc_entry_ids, ns_decls):
    """Process document.xml: fix heading levels, assign styles, remove numPr."""
    # Build style name map from styles.xml
    styles_tree = ET.parse(styles_path)
    style_name_map = build_style_name_map(styles_tree.getroot())

    tree = ET.parse(doc_path)
    root = tree.getroot()

    # body element
    body = root.find(w('body'))
    if body is None:
        print("Warning: no w:body found in document.xml", file=sys.stderr)
        return

    paragraphs = list(body.iter(w('p')))

    # ── Phase 1: Scan heading levels ──
    heading_levels_found = set()
    para_info = []  # (para_el, current_level_or_None, text)

    for para in paragraphs:
        ppr = para.find(w('pPr'))
        style_id = None
        if ppr is not None:
            pstyle = ppr.find(w('pStyle'))
            if pstyle is not None:
                style_id = pstyle.get(w('val'))

        level = get_style_level(style_id, style_name_map)
        text = get_paragraph_text(para)

        # Also check outlineLvl as backup
        if level is None and ppr is not None:
            olvl = ppr.find(w('outlineLvl'))
            if olvl is not None:
                try:
                    level = int(olvl.get(w('val'), '')) + 1  # outlineLvl is 0-based
                except ValueError:
                    pass

        para_info.append((para, level, text, style_id))
        if level is not None:
            heading_levels_found.add(level)

    # ── Phase 2: Compute level offset ──
    offset = 0
    if heading_levels_found:
        min_level = min(heading_levels_found)
        if min_level > 1:
            offset = min_level - 1

    # ── Phase 3: Apply corrections ──
    stats = {'Heading1': 0, 'Heading2': 0, 'Heading3': 0, 'TOCHeading': 0,
             'WanrenExplanation': 0, 'Normal': 0, 'toc_entry_processed': 0,
             'numPr_removed': 0, 'direct_fmt_cleared': 0}

    # Paragraphs inside tables/text boxes are layout containers, not top-level
    # flow — never let the bare-text "目录" fallback restyle those.
    in_aux = set()
    for tbl in body.iter(w('tbl')):
        in_aux.update(tbl.iter(w('p')))
    for txbx in body.iter(w('txbxContent')):
        in_aux.update(txbx.iter(w('p')))

    for para, level, text, style_id in para_info:
        # ── TOC entries: fix font/bold but preserve TOC structure ──
        # Keep the pStyle (toc N) and paragraph-level formatting (spacing,
        # jc, ind, tab stops) that give the generated TOC its indents and
        # dot leaders.  Only strip direct run-level font overrides and bold
        # so the entries inherit 宋体 / Times New Roman (and no bold) from
        # the toc-N style that update_toc_entry_styles() just fixed.
        if style_id and style_id in toc_entry_ids:
            stats['toc_entry_processed'] += 1
            for run in para.iter(w('r')):
                rpr = run.find(w('rPr'))
                if rpr is not None:
                    if rpr.find(w('rFonts')) is not None:
                        remove_child(rpr, 'rFonts')
                        stats['direct_fmt_cleared'] += 1
                    remove_child(rpr, 'b')
                    remove_child(rpr, 'bCs')
            mark_ppr = para.find(w('pPr'))
            if mark_ppr is not None:
                mark_rpr = mark_ppr.find(w('rPr'))
                if mark_rpr is not None:
                    if mark_rpr.find(w('rFonts')) is not None:
                        remove_child(mark_rpr, 'rFonts')
                    remove_child(mark_rpr, 'b')
                    remove_child(mark_rpr, 'bCs')
            continue

        ppr = para.find(w('pPr'))
        if ppr is None:
            ppr = ET.Element(w('pPr'))
            para.insert(0, ppr)

        # Remove numPr (auto-numbering); note whether it existed first — a list
        # item whose text happens to be "目录" is content, not a TOC title.
        had_numpr = ppr.find(w('numPr')) is not None
        if had_numpr:
            stats['numPr_removed'] += 1
        remove_child(ppr, 'numPr')

        # Determine target style (logical key)
        logical = None
        is_toc = False
        is_explanation = False

        # Check if it's a TOC heading: prefer the style name; the bare-text
        # "目录" fallback only fires for plain top-level paragraphs.
        if style_id and TOC_HEADING_PATTERN.match(style_name_map.get(style_id, style_id)):
            is_toc = True
        elif (text.strip() == '目录' and not had_numpr and para not in in_aux):
            is_toc = True

        if is_toc:
            logical = 'TOCHeading'
        elif text.strip().startswith('【万人教育解析】'):
            logical = 'Normal'
            is_explanation = True
        elif level is not None:
            corrected = level - offset
            if corrected < 1:
                corrected = 1
            if corrected > 3:
                # We only define up to Heading3; beyond that, treat as Normal
                logical = 'Normal'
            else:
                logical = HEADING_LOGICAL_IDS[corrected]
        else:
            # Regular paragraph → Normal
            logical = 'Normal'

        stats['WanrenExplanation' if is_explanation else logical] += 1
        target_style = resolved.get(logical, logical)

        # Set pStyle
        pstyle_el = ppr.find(w('pStyle'))
        if pstyle_el is None:
            pstyle_el = ET.Element(w('pStyle'))
            ppr.insert(0, pstyle_el)
        pstyle_el.set(w('val'), target_style)

        # Strip paragraph-level direct formatting that would override the
        # style we just assigned (spacing/alignment/indent live in the style).
        for tag in ('spacing', 'jc', 'ind'):
            remove_child(ppr, tag)

        # Run-level fonts: a direct rFonts on a run beats the style, so
        # 宋体/华文新魏 hardcoded in runs must go. Size/colour/bold are left
        # untouched elsewhere — inline emphasis is authored content, not layout.
        # Exception A: 【万人教育解析】 paragraphs use Normal (宋体) as the
        # paragraph style, but every run covering the label prefix keeps 华文新魏.
        # Word may split the label across several runs, so mark by text position
        # rather than "first run with text".
        # Exception B: the "目录" title paragraph is never bold (用户要求目录页
        # 所有字不加粗), so strip direct run-level bold on it too.
        if is_explanation:
            label_runs = mark_label_runs(para, '【万人教育解析】')
            for run, is_label in zip(para.iter(w('r')), label_runs):
                rpr = run.find(w('rPr'))
                if is_label:
                    if rpr is None:
                        rpr = ET.Element(w('rPr'))
                        run.insert(0, rpr)
                    rf = find_or_create(rpr, 'rFonts')
                    rf.set(w('eastAsia'), '华文新魏')
                elif rpr is not None and rpr.find(w('rFonts')) is not None:
                    remove_child(rpr, 'rFonts')
                    stats['direct_fmt_cleared'] += 1
        else:
            for run in para.iter(w('r')):
                rpr = run.find(w('rPr'))
                if rpr is not None and rpr.find(w('rFonts')) is not None:
                    remove_child(rpr, 'rFonts')
                    stats['direct_fmt_cleared'] += 1
                if is_toc:
                    remove_child(rpr, 'b')
                    remove_child(rpr, 'bCs')
        # The paragraph mark's own rPr carries a font too.
        mark_rpr = ppr.find(w('rPr'))
        if mark_rpr is not None:
            remove_child(mark_rpr, 'rFonts')
            if is_toc:
                remove_child(mark_rpr, 'b')
                remove_child(mark_rpr, 'bCs')

        # Update outlineLvl for headings (inserted at the schema-correct spot).
        if logical in HEADING_LOGICAL_IDS.values():
            heading_num = int(logical[-1])
            set_outline_lvl(ppr, heading_num - 1)
        else:
            remove_child(ppr, 'outlineLvl')

    # ── Phase 4: Fix page margins (header/footer distance) ──
    for sect_pr in root.iter(w('sectPr')):
        pg_mar = sect_pr.find(w('pgMar'))
        if pg_mar is not None:
            pg_mar.set(w('header'), '737')   # 1.3cm
            pg_mar.set(w('footer'), '510')   # 0.9cm

    write_tree(tree, doc_path, ns_decls)
    return offset, stats


def process_header_footer_files(unpacked_dir, resolved):
    """Apply Header/Footer style to paragraphs in header/footer XML files."""
    word_dir = os.path.join(unpacked_dir, 'word')
    touched = 0
    for fname in sorted(os.listdir(word_dir)):
        if fname.startswith('header') and fname.endswith('.xml'):
            logical = 'Header'
        elif fname.startswith('footer') and fname.endswith('.xml'):
            logical = 'Footer'
        else:
            continue

        target_style = resolved.get(logical, logical)
        fpath = os.path.join(word_dir, fname)
        ns_decls = snapshot_namespaces(fpath)
        tree = ET.parse(fpath)
        root = tree.getroot()
        for para in root.iter(w('p')):
            ppr = para.find(w('pPr'))
            if ppr is None:
                ppr = ET.Element(w('pPr'))
                para.insert(0, ppr)
            pstyle_el = ppr.find(w('pStyle'))
            if pstyle_el is None:
                pstyle_el = ET.Element(w('pStyle'))
                ppr.insert(0, pstyle_el)
            pstyle_el.set(w('val'), target_style)
            remove_child(ppr, 'jc')
            for run in para.iter(w('r')):
                rpr = run.find(w('rPr'))
                if rpr is not None:
                    remove_child(rpr, 'rFonts')
        write_tree(tree, fpath, ns_decls)
        touched += 1
    return touched


# ─── Validation ───────────────────────────────────────────────────────────────

def validate_output(path):
    """Fail loudly rather than hand back a docx Word refuses to open.

    Checks the two failure modes that actually bite: a prefix listed in
    mc:Ignorable but never declared (namespace error → "unreadable content"),
    and a pStyle pointing at a styleId that doesn't exist (silent style loss).
    """
    problems = []
    with zipfile.ZipFile(path, 'r') as zf:
        names = zf.namelist()
        if not names or names[0] != '[Content_Types].xml':
            problems.append('[Content_Types].xml is not the first zip entry')

        style_ids = set()
        if 'word/styles.xml' in names:
            sroot = ET.fromstring(zf.read('word/styles.xml'))
            style_ids = {s.get(w('styleId')) for s in sroot.iter(w('style'))}

        for name in names:
            if not name.endswith('.xml'):
                continue
            data = zf.read(name)

            # Namespace well-formedness.
            found = _find_root_start_tag(data)
            if found:
                declared = {p for p, _ in _XMLNS_RE.findall(found[2])}
                ign = re.search(r'mc:Ignorable="([^"]*)"', found[2])
                if ign:
                    undeclared = sorted(set(ign.group(1).split()) - declared)
                    if undeclared:
                        problems.append(
                            f'{name}: mc:Ignorable references undeclared prefixes {undeclared}')

            # Parse check.
            try:
                root = ET.fromstring(data)
            except ET.ParseError as exc:
                problems.append(f'{name}: XML parse error: {exc}')
                continue

            # Dangling pStyle references.
            if style_ids and name.startswith('word/'):
                missing = set()
                for ps in root.iter(w('pStyle')):
                    val = ps.get(w('val'))
                    if val and val not in style_ids:
                        missing.add(val)
                if missing:
                    problems.append(
                        f'{name}: pStyle references undefined styleIds {sorted(missing)}')

    if problems:
        raise RuntimeError(
            'Output validation failed; original file left untouched:\n  - '
            + '\n  - '.join(problems))


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='万人教育 Word 文档格式纠正')
    parser.add_argument('input', help='输入 .docx 文件路径')
    parser.add_argument('-o', '--output', help='输出 .docx 文件路径（默认覆盖原文件）')
    args = parser.parse_args()

    input_path = os.path.abspath(args.input)
    output_path = os.path.abspath(args.output) if args.output else input_path

    if not os.path.isfile(input_path):
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Create temp directory for unpacking
    tmp_dir = tempfile.mkdtemp(prefix='gank_word_')
    tmp_out = output_path + '.gank.tmp'
    try:
        # Unpack
        with zipfile.ZipFile(input_path, 'r') as zf:
            zf.extractall(tmp_dir)

        # Remove any symlinks (untrusted content)
        for dirpath, dirnames, filenames in os.walk(tmp_dir):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if os.path.islink(fp):
                    os.unlink(fp)

        styles_path = os.path.join(tmp_dir, 'word', 'styles.xml')
        doc_path = os.path.join(tmp_dir, 'word', 'document.xml')

        if not os.path.isfile(styles_path):
            print("Error: word/styles.xml not found in docx", file=sys.stderr)
            sys.exit(1)
        if not os.path.isfile(doc_path):
            print("Error: word/document.xml not found in docx", file=sys.stderr)
            sys.exit(1)

        # Process
        styles_ns = snapshot_namespaces(styles_path)
        doc_ns = snapshot_namespaces(doc_path)

        resolved, toc_entry_ids = update_styles_xml(styles_path, styles_ns)
        offset, stats = process_document_xml(
            doc_path, styles_path, resolved, toc_entry_ids, doc_ns)
        hf_count = process_header_footer_files(tmp_dir, resolved)

        # Repack, preserving the original part order (Word expects
        # [Content_Types].xml first) and keeping stored/deflated modes.
        with zipfile.ZipFile(input_path, 'r') as zf:
            original_order = zf.namelist()
            compress_types = {i.filename: i.compress_type for i in zf.infolist()}

        written = set()
        with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zf:
            for arcname in original_order:
                filepath = os.path.join(tmp_dir, arcname)
                if not os.path.isfile(filepath):
                    continue
                zf.write(filepath, arcname,
                         compress_type=compress_types.get(arcname, zipfile.ZIP_DEFLATED))
                written.add(arcname)
            # Any part we created that wasn't in the original.
            for dirpath, _dirnames, filenames in os.walk(tmp_dir):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    arcname = os.path.relpath(filepath, tmp_dir)
                    if arcname not in written:
                        zf.write(filepath, arcname)

        validate_output(tmp_out)
        os.replace(tmp_out, output_path)

        # ── Report ──
        print(f"Done: {output_path}")
        print(f"  标题层级偏移      : {offset} (0 = 原层级已从一级开始)")
        print(f"  一级标题          : {stats['Heading1']}  → styleId {resolved['Heading1']!r}")
        print(f"  二级标题          : {stats['Heading2']}  → styleId {resolved['Heading2']!r}")
        print(f"  三级标题          : {stats['Heading3']}  → styleId {resolved['Heading3']!r}")
        print(f"  目录标题          : {stats['TOCHeading']}  → styleId {resolved['TOCHeading']!r}")
        print(f"  解析段落          : {stats['WanrenExplanation']}  → styleId {resolved['Normal']!r}（标签 run 为华文新魏，正文继承宋体）")
        print(f"  正文段落          : {stats['Normal']}  → styleId {resolved['Normal']!r}")
        print(f"  目录条目(字体修正)  : {stats['toc_entry_processed']}")
        print(f"  清除自动编号      : {stats['numPr_removed']}")
        print(f"  清除直接字体      : {stats['direct_fmt_cleared']}")
        print(f"  页眉页脚部件      : {hf_count}")

    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        if os.path.exists(tmp_out):
            os.unlink(tmp_out)


if __name__ == '__main__':
    main()
