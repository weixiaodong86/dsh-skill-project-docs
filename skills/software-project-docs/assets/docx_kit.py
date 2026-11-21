# -*- coding: utf-8 -*-
"""docx_kit —— 软件研发项目管理过程文档的 DOCX 构建库。

把受控 Markdown 内容源渲染为符合统一排版规范的正式 Word 文档：
封面、修订记录、自动目录、分节页码、页眉、标题自动编号、
三线表 / 全框线表、图与图表编号、模板填写说明（楷体斜体）。

典型用法：
    from docx_kit import build_document
    build_document("content.md", "out.docx")

依赖：python-docx（随 DSH 运行时提供）。
"""

from __future__ import annotations

import os
import re
from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------------------
# 版式常量（改这里即可整体换版式）
# ---------------------------------------------------------------------------
PAGE = {
    "top": 2.54,
    "bottom": 2.54,
    "left": 3.0,
    "right": 2.5,
    "header": 1.5,
    "footer": 1.75,
    "width_cm": 21.0,
    "height_cm": 29.7,
}
CONTENT_WIDTH_CM = PAGE["width_cm"] - PAGE["left"] - PAGE["right"]  # 15.5cm

FONT_BODY = "宋体"
FONT_HEAD = "黑体"
FONT_KAI = "楷体"
FONT_LATIN = "Times New Roman"

SIZE_BODY = 12.0        # 小四
SIZE_H1 = 16.0          # 三号
SIZE_H2 = 14.0          # 四号
SIZE_H3 = 12.0          # 小四
SIZE_CAPTION = 10.5     # 五号
SIZE_TABLE = 9.0        # 小五
SIZE_NOTE = 10.5
SIZE_HDR = 9.0
SIZE_COVER_TITLE = 26.0  # 一号
SIZE_COVER_PROJECT = 22.0  # 二号

LINE_SPACING = 1.5
COLOR_ANNOTATION = "595959"
COLOR_HEADER_FILL = "F2F2F2"

STYLE_BODY = "PM正文"
STYLE_TABLE = "PM表格"
STYLE_TABLE_CAPTION = "PM表题"
STYLE_FIGURE_CAPTION = "PM图题"
STYLE_FILL = "PM填写说明"
STYLE_EXAMPLE = "PM示例"
STYLE_NOTE = "PM注解"
STYLE_NUMBERED = "PM条目"

_BLANK = "\u3000"  # 全角空格，用于空白占位段落


# ---------------------------------------------------------------------------
# 低层 XML 工具
# ---------------------------------------------------------------------------
def _set_rfonts(rpr, latin: str, eastasia: str) -> None:
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), eastasia)
    rfonts.set(qn("w:cs"), latin)


def _font(style_or_run, latin=FONT_BODY, eastasia=FONT_BODY, size=SIZE_BODY,
          bold=False, italic=False, color=None) -> None:
    font = style_or_run.font
    font.name = latin
    font.size = Pt(size)
    font.bold = bold
    font.italic = italic
    if color:
        font.color.rgb = RGBColor.from_string(color)
    else:
        font.color.rgb = RGBColor.from_string("000000")
    el = getattr(style_or_run, "element", None)
    if el is None:
        el = style_or_run._element
    rpr = el.get_or_add_rPr()
    _set_rfonts(rpr, latin, eastasia)


def _insert(parent, el, successors) -> None:
    """按 OOXML schema 顺序插入元素，避免 Word 报“内容有问题”。"""
    for tag in successors:
        found = parent.find(qn(tag))
        if found is not None:
            found.addprevious(el)
            return
    parent.append(el)


def _ind(ppr, first_line_chars=0, size_pt=SIZE_BODY,
         left_cm=None, hanging_cm=None) -> None:
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        _insert(ppr, ind, ("w:contextualSpacing", "w:mirrorIndents",
                           "w:suppressOverlap", "w:jc", "w:textDirection",
                           "w:textAlignment", "w:textboxTightWrap",
                           "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
                           "w:sectPr", "w:pPrChange"))
    if first_line_chars:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        ind.set(qn("w:firstLine"), str(int(size_pt * first_line_chars / 100 * 20)))
    if left_cm is not None:
        ind.set(qn("w:left"), str(int(left_cm * 567)))
        ind.set(qn("w:leftChars"), "0")
    if hanging_cm is not None:
        ind.set(qn("w:hanging"), str(int(hanging_cm * 567)))
        ind.set(qn("w:hangingChars"), "0")


def _outline_level(style, level: int) -> None:
    ppr = style.element.get_or_add_pPr()
    for old in ppr.findall(qn("w:outlineLvl")):
        ppr.remove(old)
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), str(level))
    ppr.append(el)


def _bottom_border(paragraph, size=6, color="000000") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    _insert(ppr, pbdr, ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                        "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                        "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
                        "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                        "w:contextualSpacing", "w:mirrorIndents",
                        "w:suppressOverlap", "w:jc", "w:textDirection",
                        "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                        "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                        "w:pPrChange"))


def _shade(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    _insert(tcpr, shd, ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                        "w:vAlign", "w:hideMark", "w:headers", "w:cellIns",
                        "w:cellDel", "w:cellMerge", "w:tcPrChange"))


def _cell_valign(cell, val="center") -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    v = OxmlElement("w:vAlign")
    v.set(qn("w:val"), val)
    _insert(tcpr, v, ("w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
                      "w:cellMerge", "w:tcPrChange"))


def _table_borders(table, style="all", size=4, head_size=8, color="000000") -> None:
    tblpr = table._tbl.tblPr
    for old in tblpr.findall(qn("w:tblBorders")):
        tblpr.remove(old)
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right", "insideH", "insideV"]
    for edge in edges:
        el = OxmlElement(f"w:{edge}")
        if style == "none":
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        elif style == "three-line" and edge in ("left", "right", "insideV", "insideH"):
            el.set(qn("w:val"), "none")
            el.set(qn("w:sz"), "0")
        else:
            el.set(qn("w:val"), "single")
            if style == "three-line" and edge in ("top", "bottom"):
                el.set(qn("w:sz"), str(head_size))
            else:
                el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    _insert(tblpr, borders, ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                             "w:tblCaption", "w:tblDescription", "w:tblPrChange"))
    if style == "three-line":
        # 表头下框线加粗
        for cell in table.rows[0].cells:
            tcpr = cell._tc.get_or_add_tcPr()
            tcb = OxmlElement("w:tcBorders")
            b = OxmlElement("w:bottom")
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), str(head_size))
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), color)
            tcb.append(b)
            _insert(tcpr, tcb, ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                                "w:tcFitText", "w:vAlign", "w:hideMark",
                                "w:headers", "w:cellIns", "w:cellDel",
                                "w:cellMerge", "w:tcPrChange"))


def _table_width(table, widths_cm) -> None:
    """同时写 tblW 与 tblGrid，保证固定布局下自定义列宽生效。"""
    tbl = table._tbl
    tblpr = tbl.tblPr
    total = sum(widths_cm)
    tw = tblpr.find(qn("w:tblW"))
    if tw is None:
        tw = OxmlElement("w:tblW")
        _insert(tblpr, tw, ("w:jc", "w:tblCellSpacing", "w:tblInd",
                            "w:tblBorders", "w:shd", "w:tblLayout",
                            "w:tblCellMar", "w:tblLook", "w:tblCaption",
                            "w:tblDescription", "w:tblPrChange"))
    tw.set(qn("w:w"), str(int(total * 567)))
    tw.set(qn("w:type"), "dxa")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return
    cols = grid.findall(qn("w:gridCol"))
    for col, width in zip(cols, widths_cm):
        col.set(qn("w:w"), str(int(width * 567)))


def _table_cell_margins(table, left_cm=0.1, right_cm=0.1, top_cm=0.05, bottom_cm=0.05) -> None:
    tblpr = table._tbl.tblPr
    for old in tblpr.findall(qn("w:tblCellMar")):
        tblpr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, cm in (("top", top_cm), ("left", left_cm),
                     ("bottom", bottom_cm), ("right", right_cm)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(cm * 567)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _insert(tblpr, mar, ("w:tblLook", "w:tblCaption", "w:tblDescription",
                         "w:tblPrChange"))


def _repeat_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def _row_height(row, cm=0.6) -> None:
    trpr = row._tr.get_or_add_trPr()
    h = OxmlElement("w:trHeight")
    h.set(qn("w:val"), str(int(cm * 567)))
    h.set(qn("w:hRule"), "atLeast")
    trpr.append(h)


def _field(paragraph, instr: str, placeholder: str = "", size=SIZE_HDR) -> None:
    parts = [
        ("begin", None),
        ("instr", instr),
        ("separate", None),
        ("text", placeholder),
        ("end", None),
    ]
    for kind, value in parts:
        run = paragraph.add_run()
        run.font.size = Pt(size)
        run.font.name = FONT_LATIN
        _set_rfonts(run._r.get_or_add_rPr(), FONT_LATIN, FONT_BODY)
        if kind == "begin":
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), "begin")
        elif kind == "separate":
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), "separate")
        elif kind == "end":
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), "end")
        elif kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {value} "
        else:
            el = OxmlElement("w:t")
            el.text = value
        run._r.append(el)


def _page_number_type(section, fmt: str, start: int) -> None:
    sectpr = section._sectPr
    for old in sectpr.findall(qn("w:pgNumType")):
        sectpr.remove(old)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:fmt"), fmt)
    el.set(qn("w:start"), str(start))
    _insert(sectpr, el, ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
                         "w:titlePg", "w:textDirection", "w:bidi",
                         "w:rtlGutter", "w:docGrid", "w:printerSettings",
                         "w:sectPrChange"))


def _setup_page(doc) -> None:
    sec = doc.sections[0]
    sec.page_width = Cm(PAGE["width_cm"])
    sec.page_height = Cm(PAGE["height_cm"])
    sec.top_margin = Cm(PAGE["top"])
    sec.bottom_margin = Cm(PAGE["bottom"])
    sec.left_margin = Cm(PAGE["left"])
    sec.right_margin = Cm(PAGE["right"])
    sec.header_distance = Cm(PAGE["header"])
    sec.footer_distance = Cm(PAGE["footer"])


def _style(doc, name, base="Normal"):
    try:
        return doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = doc.styles[base]
        return style


def _setup_styles(doc) -> None:
    normal = doc.styles["Normal"]
    _font(normal, FONT_LATIN, FONT_BODY, SIZE_BODY)
    pf = normal.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.widow_control = True

    # 标题：黑体、黑色、不缩进、与下段同页
    heads = [
        ("Heading 1", SIZE_H1, 18, 12),
        ("Heading 2", SIZE_H2, 12, 6),
        ("Heading 3", SIZE_H3, 6, 6),
    ]
    for name, size, before, after in heads:
        st = doc.styles[name]
        _font(st, FONT_LATIN, FONT_HEAD, size, bold=True)
        pf = st.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = 1.25
        pf.keep_with_next = True
        pf.first_line_indent = Pt(0)

    try:
        st4 = doc.styles["Heading 4"]
        _font(st4, FONT_LATIN, FONT_HEAD, SIZE_H3, bold=True)
    except KeyError:
        pass

    body = _style(doc, STYLE_BODY)
    _font(body, FONT_LATIN, FONT_BODY, SIZE_BODY)
    pf = body.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    _ind(body.element.get_or_add_pPr(), first_line_chars=200, size_pt=SIZE_BODY)

    note = _style(doc, STYLE_NOTE)
    _font(note, FONT_LATIN, FONT_BODY, SIZE_NOTE)
    pf = note.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = 1.25
    _ind(note.element.get_or_add_pPr(), first_line_chars=200, size_pt=SIZE_NOTE)

    fill = _style(doc, STYLE_FILL)
    _font(fill, FONT_LATIN, FONT_KAI, SIZE_BODY, italic=True, color=COLOR_ANNOTATION)
    pf = fill.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    _ind(fill.element.get_or_add_pPr(), first_line_chars=200, size_pt=SIZE_BODY)

    example = _style(doc, STYLE_EXAMPLE)
    _font(example, FONT_LATIN, FONT_KAI, SIZE_BODY, italic=True, color=COLOR_ANNOTATION)
    pf = example.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    _ind(example.element.get_or_add_pPr(), first_line_chars=200, size_pt=SIZE_BODY)

    cap = _style(doc, STYLE_TABLE_CAPTION)
    _font(cap, FONT_LATIN, FONT_HEAD, SIZE_CAPTION, bold=True)
    pf = cap.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.line_spacing = 1.0
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    pf.keep_with_next = True
    pf.first_line_indent = Pt(0)

    figcap = _style(doc, STYLE_FIGURE_CAPTION)
    _font(figcap, FONT_LATIN, FONT_HEAD, SIZE_CAPTION, bold=True)
    pf = figcap.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.line_spacing = 1.0
    pf.space_before = Pt(3)
    pf.space_after = Pt(6)
    pf.first_line_indent = Pt(0)

    cell = _style(doc, STYLE_TABLE)
    _font(cell, FONT_LATIN, FONT_BODY, SIZE_TABLE)
    pf = cell.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.line_spacing = 1.0
    pf.space_before = Pt(1)
    pf.space_after = Pt(1)
    pf.first_line_indent = Pt(0)

    numbered = _style(doc, STYLE_NUMBERED)
    _font(numbered, FONT_LATIN, FONT_BODY, SIZE_BODY)
    pf = numbered.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    _ind(numbered.element.get_or_add_pPr(), left_cm=0.74, hanging_cm=0.74,
         size_pt=SIZE_BODY)


# ---------------------------------------------------------------------------
# 封面 / 前置部分 / 页眉页脚
# ---------------------------------------------------------------------------
def _run(p, text, latin=FONT_LATIN, eastasia=FONT_BODY, size=SIZE_BODY,
         bold=False, italic=False, color=None):
    r = p.add_run(text)
    r.font.name = latin
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = RGBColor.from_string(color or "000000")
    _set_rfonts(r._element.get_or_add_rPr(), latin, eastasia)
    return r


def _blank(doc, size=SIZE_BODY, count=1):
    for _ in range(count):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        _run(p, _BLANK, size=size)


def _date_cn(raw: str) -> str:
    s = (raw or "").strip()
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2)):02d}月{int(m.group(3)):02d}日"
    return s or date.today().strftime("%Y年%m月%d日")


def _cover(doc, meta) -> None:
    for _ in range(1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _run(p, meta.get("classification", "内部"), size=10.5, eastasia=FONT_HEAD, bold=True)

    _blank(doc, count=4)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, meta.get("project", ""), size=SIZE_COVER_PROJECT, eastasia=FONT_HEAD, bold=True)

    _blank(doc, count=2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    _run(p, meta.get("title", "项目管理文档"), size=SIZE_COVER_TITLE,
         eastasia=FONT_HEAD, bold=True)

    _blank(doc, count=1)
    for label, value in (
        ("文档编号", meta.get("doc_code", "")),
        ("版 本 号", meta.get("version", "V1.0")),
    ):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, f"{label}：{value}", size=SIZE_BODY)

    _blank(doc, count=6)

    sign = doc.add_table(rows=4, cols=4)
    sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign.autofit = False
    _table_borders(sign, style="none")
    _table_cell_margins(sign, 0.3, 0.3, 0.15, 0.15)
    roles = [
        ("身 份", "姓 名", "签 字", "日 期"),
        ("编 制", meta.get("author", ""), "", ""),
        ("审 核", meta.get("reviewer", ""), "", ""),
        ("批 准", meta.get("approver", ""), "", ""),
    ]
    widths = [2.6, 3.6, 4.6, 3.4]
    _table_width(sign, widths)
    for i, values in enumerate(roles):
        for j, cell in enumerate(sign.rows[i].cells):
            cell.width = Cm(widths[j])
            cell.text = ""
            para = cell.paragraphs[0]
            para.style = doc.styles[STYLE_TABLE]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(para, values[j], size=SIZE_BODY,
                 eastasia=FONT_HEAD if (i == 0 or j == 0) else FONT_BODY,
                 bold=(i == 0))

    _blank(doc, count=3)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, meta.get("issuer", ""), size=SIZE_BODY)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, _date_cn(meta.get("date", "")), size=SIZE_BODY)


def _header(section, meta) -> None:
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.text = ""
    p.paragraph_format.tab_stops.add_tab_stop(
        Cm(CONTENT_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT)
    left = "  ".join(x for x in [meta.get("project", ""), meta.get("title", "")] if x)
    right = "  ".join(x for x in [meta.get("doc_code", ""), meta.get("version", "")] if x)
    _run(p, left, size=SIZE_HDR)
    _run(p, "\t", size=SIZE_HDR)
    _run(p, right, size=SIZE_HDR)
    _bottom_border(p, size=6)


def _footer(section, with_page=True) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if not with_page:
        return
    _run(p, "第 ", size=SIZE_HDR)
    _field(p, "PAGE", "1", SIZE_HDR)
    _run(p, " 页  共 ", size=SIZE_HDR)
    _field(p, "NUMPAGES", "1", SIZE_HDR)
    _run(p, " 页", size=SIZE_HDR)


def _revision_table(doc, meta, heading_size=SIZE_H1) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6 if heading_size < SIZE_H1 else 12)
    _run(p, "修  订  记  录", size=heading_size, eastasia=FONT_HEAD, bold=True)

    revisions = meta.get("revisions") or []
    rows = [["版本号", "日期", "修订人", "修订说明", "审核人"]]
    if not revisions:
        rows.append([meta.get("version", "V1.0"), meta.get("date", ""),
                     meta.get("author", ""), "首次发布", meta.get("reviewer", "")])
    for item in revisions:
        item = list(item) + [""] * (5 - len(item))
        rows.append(item[:5])

    table = doc.add_table(rows=len(rows), cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _table_borders(table, style="all", size=4, head_size=8)
    _table_cell_margins(table, 0.15, 0.15, 0.06, 0.06)
    widths = [2.2, 2.6, 2.0, 6.2, 2.5]
    _table_width(table, widths)
    for i, row in enumerate(rows):
        _row_height(table.rows[i], 0.7)
        for j, value in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.width = Cm(widths[j])
            cell.text = ""
            para = cell.paragraphs[0]
            para.style = doc.styles[STYLE_TABLE]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _run(para, str(value), size=SIZE_TABLE, bold=(i == 0))
            if i == 0:
                _shade(cell, COLOR_HEADER_FILL)
            _cell_valign(cell, "center")
    _repeat_header(table.rows[0])


def _toc(doc) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    _run(p, "目      录", size=SIZE_H1, eastasia=FONT_HEAD, bold=True)

    para = doc.add_paragraph()
    para.paragraph_format.first_line_indent = Pt(0)
    para.paragraph_format.line_spacing = 1.25
    _field(para, r'TOC \o "1-3" \h \z \u', "")


def _brief_head(doc, meta) -> None:
    """短文档版式：不设封面与目录，标题块 + 修订记录 + 正文。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    _run(p, meta.get("title", "项目管理文档"), size=18, eastasia=FONT_HEAD, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    line = "文档编号：{code}　　版本：{ver}　　日期：{date}".format(
        code=meta.get("doc_code", ""),
        ver=meta.get("version", "V1.0"),
        date=_date_cn(meta.get("date", "")),
    )
    _run(p, line, size=SIZE_HDR)

    _revision_table(doc, meta, heading_size=SIZE_BODY)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    _bottom_border(p, size=6)


# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------
def parse_front_matter(text: str):
    if not text.startswith("---"):
        return {}, text
    end = text.find("\n---", 3)
    if end < 0:
        return {}, text
    raw = text[3:end].strip("\n")
    body = text[end + 4:]
    meta = {}
    revisions = []
    current = None
    for line in raw.splitlines():
        if not line.strip() or line.strip().startswith("#"):
            continue
        if line.lstrip().startswith("- "):
            if current == "revisions":
                revisions.append([x.strip() for x in line.lstrip()[2:].split("|")])
            continue
        if ":" in line and not line.startswith((" ", "\t")):
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if value == "":
                current = key
                meta[key] = [] if key == "revisions" else ""
            else:
                current = None
                meta[key] = value
    if revisions:
        meta["revisions"] = revisions
    return meta, body


_HEADING_NUM = re.compile(r"^\d+(?:\.\d+)*\.?\s+|^第[一二三四五六七八九十]+[章节]\s*")
_APPENDIX = re.compile(r"^附录\s*([A-Za-z一二三四五六七八九十]+)")


def _strip_number(text: str) -> str:
    return _HEADING_NUM.sub("", text.strip())


def _inline_runs(paragraph, text, style_size=SIZE_BODY, font=FONT_BODY,
                 italic=False, color=None, bold=False) -> None:
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not part:
            continue
        _run(paragraph, part, size=style_size, eastasia=font,
             bold=bold or (i % 2 == 1), italic=italic, color=color)


def parse_blocks(body: str):
    blocks = []
    lines = body.splitlines()
    i = 0
    n = len(lines)
    pending_caption = None  # ("table"|"figure", caption, extra)
    while i < n:
        raw = lines[i].rstrip()
        s = raw.strip()
        if not s:
            i += 1
            continue

        if s.startswith("```"):
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith("```"):
                buf.append(lines[i].rstrip())
                i += 1
            i += 1
            blocks.append(("code", "\n".join(buf)))
            continue

        if s.startswith("#"):
            level = len(s) - len(s.lstrip("#"))
            blocks.append(("h", min(level, 4), s[level:].strip()))
            i += 1
            continue

        if s.startswith(":::"):
            head = s[3:].strip()
            kind = head.split()[0] if head else "note"
            rest = head[len(kind):].strip()
            i += 1
            buf = [rest] if rest else []
            while i < n and not lines[i].strip().startswith(":::"):
                buf.append(lines[i].rstrip())
                i += 1
            i += 1
            blocks.append((kind if kind in ("fill", "example", "note") else "note",
                           "\n".join(buf).strip("\n")))
            continue

        cap_match = re.match(r"^表\s*[:：]\s*(.*)$", s)
        if cap_match:
            cap = cap_match.group(1).strip()
            widths = None
            if "||" in cap:
                cap, w = cap.split("||", 1)
                widths = [float(x) for x in w.split(",") if x.strip()]
            pending_caption = ("table", cap.strip(), widths)
            i += 1
            continue

        fig_match = re.match(r"^图\s*[:：]\s*(.*)$", s)
        if fig_match:
            cap = fig_match.group(1).strip()
            path, width = "", None
            if "||" in cap:
                parts = [x.strip() for x in cap.split("||")]
                cap = parts[0]
                if len(parts) > 1:
                    path = parts[1]
                if len(parts) > 2 and parts[2]:
                    width = float(parts[2].replace("cm", ""))
            blocks.append(("figure", cap.strip(), path, width))
            i += 1
            continue

        if s.startswith("|"):
            rows = []
            while i < n and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            if rows:
                width_spec = None
                if pending_caption and pending_caption[0] == "table":
                    width_spec = pending_caption[2]
                    cap = pending_caption[1]
                    pending_caption = None
                else:
                    cap = None
                blocks.append(("table", cap, width_spec, rows))
            continue

        if s.startswith("!["):
            m = re.match(r"^!\[(.*?)\]\((.*?)\)(?:\{width=([\d.]+)cm\})?$", s)
            if m:
                blocks.append(("figure", m.group(1), m.group(2),
                               float(m.group(3)) if m.group(3) else None))
                i += 1
                continue

        if s.startswith("- ") or s.startswith("* "):
            items = []
            while i < n and re.match(r"^\s*[-*]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*]\s+", "", lines[i]).strip())
                i += 1
            blocks.append(("ul", items))
            continue

        if re.match(r"^\d+[.)]\s+", s) or s.startswith("（") and "）" in s:
            items = []
            while i < n and (re.match(r"^\s*\d+[.)]\s+", lines[i])
                             or re.match(r"^\s*（\d+）", lines[i])):
                item = re.sub(r"^\s*\d+[.)]\s+", "", lines[i]).strip()
                item = re.sub(r"^\s*（\d+）\s*", "", item)
                items.append(item)
                i += 1
            blocks.append(("ol", items))
            continue

        # 普通段落（连续行合并）
        buf = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(("#", "|", ":::", "- ", "* ", "```"))
                    or re.match(r"^(表|图)\s*[:：]", nxt)
                    or re.match(r"^\d+[.)]\s+", nxt)
                    or nxt.startswith("![")):
                break
            buf.append(nxt)
            i += 1
        blocks.append(("p", "".join(buf) if _is_cjk_wrap(buf) else "\n".join(buf)))

    # 悬空的表格题：生成表题但不生成表格，交由自检报错，避免内容被静默丢弃
    if pending_caption and pending_caption[0] == "table":
        blocks.append(("dangling", pending_caption[1]))
    return blocks


def _is_cjk_wrap(lines) -> bool:
    return all(re.search(r"[\u4e00-\u9fff]$", ln) for ln in lines[:-1])


# ---------------------------------------------------------------------------
# 渲染
# ---------------------------------------------------------------------------
class _Renderer:
    def __init__(self, doc, meta):
        self.doc = doc
        self.meta = meta
        self.counters = [0, 0, 0, 0]
        self.chapter = "1"
        self.appendix = False
        self.tbl_no = 0
        self.fig_no = 0

    def heading(self, level: int, text: str) -> None:
        text = _strip_number(text)
        app = _APPENDIX.match(text)
        if app and level == 1:
            label = app.group(1).upper()
            self.chapter = label
            self.appendix = True
            self.counters = [0, 0, 0, 0]
            self.tbl_no = 0
            self.fig_no = 0
            display = text
        elif level == 1:
            self.counters[0] += 1
            self.counters[1:] = [0, 0, 0]
            self.chapter = str(self.counters[0])
            self.appendix = False
            self.tbl_no = 0
            self.fig_no = 0
            display = f"{self.counters[0]} {text}"
        else:
            idx = level - 1
            self.counters[idx] += 1
            for k in range(idx + 1, 4):
                self.counters[k] = 0
            if self.appendix:
                prefix = self.chapter + "." + ".".join(
                    str(c) for c in self.counters[1:level])
            else:
                prefix = ".".join(str(c) for c in self.counters[:level])
            display = f"{prefix} {text}"

        style = "Heading 1" if level == 1 else ("Heading 2" if level == 2 else
                                                 ("Heading 3" if level == 3 else "Heading 4"))
        try:
            p = self.doc.add_paragraph(display, style=style)
        except KeyError:
            p = self.doc.add_paragraph(display, style="Heading 3")
        p.paragraph_format.first_line_indent = Pt(0)

    def paragraph(self, text: str) -> None:
        p = self.doc.add_paragraph(style=STYLE_BODY)
        _inline_runs(p, text)

    def bullets(self, items, ordered=False) -> None:
        for idx, item in enumerate(items, 1):
            p = self.doc.add_paragraph(style=STYLE_NUMBERED)
            prefix = f"（{idx}）" if ordered else "●  "
            _run(p, prefix, size=SIZE_BODY)
            _inline_runs(p, item)

    def table(self, caption, width_spec, rows) -> None:
        if caption:
            self.tbl_no += 1
            p = self.doc.add_paragraph(style=STYLE_TABLE_CAPTION)
            _run(p, f"表{self.chapter}-{self.tbl_no}  {caption}",
                 size=SIZE_CAPTION, eastasia=FONT_HEAD, bold=True)
        cols = max(len(r) for r in rows)
        table = self.doc.add_table(rows=len(rows), cols=cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        _table_borders(table, style="all", size=4, head_size=8)
        _table_cell_margins(table, 0.15, 0.15, 0.06, 0.06)

        if width_spec and len(width_spec) == cols:
            total = sum(width_spec)
            widths_cm = [CONTENT_WIDTH_CM * w / total for w in width_spec]
        else:
            widths_cm = [CONTENT_WIDTH_CM / cols] * cols
        _table_width(table, widths_cm)
        widths = [Cm(w) for w in widths_cm]

        for i, row in enumerate(rows):
            _row_height(table.rows[i], 0.62)
            for j in range(cols):
                value = row[j] if j < len(row) else ""
                cell = table.rows[i].cells[j]
                cell.width = widths[j]
                cell.text = ""
                para = cell.paragraphs[0]
                para.style = self.doc.styles[STYLE_TABLE]
                center = (i == 0) or len(value) <= 16 or j > 0 and len(value) <= 24
                para.alignment = (WD_ALIGN_PARAGRAPH.CENTER if center
                                  else WD_ALIGN_PARAGRAPH.LEFT)
                _inline_runs(para, str(value), style_size=SIZE_TABLE, bold=(i == 0))
                if i == 0:
                    _shade(cell, COLOR_HEADER_FILL)
                _cell_valign(cell, "center")
        _repeat_header(table.rows[0])

    def figure(self, caption, path, width_cm=None) -> None:
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run()
        width = Cm(width_cm) if width_cm else Cm(min(14.5, CONTENT_WIDTH_CM))
        if path and os.path.exists(path):
            run.add_picture(path, width=width)
        else:
            _run(p, f"［图片缺失：{path or '未提供'}］", size=SIZE_BODY,
                 eastasia=FONT_KAI, italic=True, color="C00000")
        self.fig_no += 1
        cap = self.doc.add_paragraph(style=STYLE_FIGURE_CAPTION)
        _run(cap, f"图{self.chapter}-{self.fig_no}  {caption or ''}",
             size=SIZE_CAPTION, eastasia=FONT_HEAD, bold=True)

    def fill(self, text: str) -> None:
        first, *rest = text.splitlines()
        p = self.doc.add_paragraph(style=STYLE_FILL)
        _run(p, "【填写说明】", size=SIZE_BODY, eastasia=FONT_KAI, italic=True,
             color=COLOR_ANNOTATION, bold=True)
        _run(p, first, size=SIZE_BODY, eastasia=FONT_KAI, italic=True,
             color=COLOR_ANNOTATION)
        for line in rest:
            p2 = self.doc.add_paragraph(style=STYLE_FILL)
            _run(p2, line, size=SIZE_BODY, eastasia=FONT_KAI, italic=True,
                 color=COLOR_ANNOTATION)

    def example(self, text: str) -> None:
        lines = text.splitlines()
        for idx, line in enumerate(lines):
            p = self.doc.add_paragraph(style=STYLE_EXAMPLE)
            if idx == 0:
                _run(p, "示例：", size=SIZE_BODY, eastasia=FONT_KAI, italic=True,
                     color=COLOR_ANNOTATION, bold=True)
            _run(p, line, size=SIZE_BODY, eastasia=FONT_KAI, italic=True,
                 color=COLOR_ANNOTATION)

    def note(self, text: str) -> None:
        p = self.doc.add_paragraph(style=STYLE_NOTE)
        _run(p, "注：", size=SIZE_NOTE, eastasia=FONT_HEAD, bold=True)
        _inline_runs(p, text, style_size=SIZE_NOTE)

    def code(self, text: str) -> None:
        for line in text.splitlines() or [""]:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1.0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            _run(p, line or _BLANK, latin="Consolas", eastasia=FONT_BODY, size=10)

    def render(self, blocks) -> None:
        for block in blocks:
            kind = block[0]
            if kind == "h":
                self.heading(block[1], block[2])
            elif kind == "p":
                self.paragraph(block[1])
            elif kind == "ul":
                self.bullets(block[1], ordered=False)
            elif kind == "ol":
                self.bullets(block[1], ordered=True)
            elif kind == "table":
                self.table(block[1], block[2], block[3])
            elif kind == "figure":
                self.figure(block[1], block[2], block[3])
            elif kind == "fill":
                self.fill(block[1])
            elif kind == "example":
                self.example(block[1])
            elif kind == "note":
                self.note(block[1])
            elif kind == "code":
                self.code(block[1])
            elif kind == "dangling":
                self.tbl_no += 1
                p = self.doc.add_paragraph(style=STYLE_TABLE_CAPTION)
                _run(p, f"表{self.chapter}-{self.tbl_no}  {block[1]}",
                     size=SIZE_CAPTION, eastasia=FONT_HEAD, bold=True)


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def build_document(md_path: str, out_path: str, meta_overrides: dict | None = None) -> dict:
    with open(md_path, "r", encoding="utf-8-sig") as fh:
        text = fh.read()
    meta, body = parse_front_matter(text)
    if meta_overrides:
        meta.update({k: v for k, v in meta_overrides.items() if v is not None})
    blocks = parse_blocks(body)

    doc = Document()
    _setup_page(doc)
    _setup_styles(doc)

    brief = str(meta.get("layout", "")).strip().lower() == "brief"
    if brief:
        main = doc.sections[0]
        _page_number_type(main, "decimal", 1)
        _header(main, meta)
        _footer(main, with_page=True)
        _brief_head(doc, meta)
    else:
        _cover(doc, meta)

        front = doc.add_section(WD_SECTION.NEW_PAGE)
        _page_number_type(front, "lowerRoman", 1)
        _header(front, meta)
        _footer(front, with_page=True)
        _revision_table(doc, meta)
        doc.add_page_break()
        _toc(doc)

        main = doc.add_section(WD_SECTION.NEW_PAGE)
        _page_number_type(main, "decimal", 1)
        _header(main, meta)
        _footer(main, with_page=True)

    renderer = _Renderer(doc, meta)
    renderer.render(blocks)

    # 打开文档时自动更新目录与页码域
    settings = doc.settings.element
    for old in settings.findall(qn("w:updateFields")):
        settings.remove(old)
    upd = OxmlElement("w:updateFields")
    upd.set(qn("w:val"), "true")
    _insert(settings, upd, (
        "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat",
        "w:docVars", "w:rsids", "w:mathPr", "w:attachedSchema",
        "w:themeFontLang", "w:clrSchemeMapping",
        "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
        "w:forceUpgrade", "w:captions", "w:readModeInkLockDown",
        "w:smartTagType", "w:schemaLibrary", "w:shapeDefaults",
        "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator"))

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(out_path)
    return {
        "path": out_path,
        "title": meta.get("title", ""),
        "version": meta.get("version", ""),
        "sections": len(doc.sections),
        "paragraphs": len(doc.paragraphs),
        "tables": len(doc.tables),
        "headings": [b[2] for b in blocks if b[0] == "h"],
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("用法: python docx_kit.py <content.md> <out.docx>")
        raise SystemExit(2)
    info = build_document(sys.argv[1], sys.argv[2])
    print(info)
