# -*- coding: utf-8 -*-
"""规范自检：检查生成的 DOCX 是否满足排版与内容规范。

用法：
    python -B scripts/check_docx.py <file.docx> [--json]

退出码：0 = 通过（可有 warning），1 = 存在 error。
"""

from __future__ import annotations

import argparse
import json
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

CONTENT_WIDTH_CM = 15.5
ALLOWED_STYLES = {
    "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4",
    "PM正文", "PM表格", "PM表题", "PM图题", "PM填写说明", "PM示例",
    "PM注解", "PM条目",
}
FONT_STYLES = {
    "PM正文": "宋体",
    "PM表格": "宋体",
    "PM表题": "黑体",
    "PM图题": "黑体",
    "PM填写说明": "楷体",
    "PM示例": "楷体",
    "PM注解": "宋体",
}
BANNED = [
    r"本次修改", r"此次修改", r"本版修改", r"本次修订", r"此次修订",
    r"根据您的要求", r"根据你的要求", r"已按您", r"已按你",
    r"AI\s*生成", r"由\s*AI", r"人工智能生成", r"以下为修改",
    r"我已完成", r"修改说明如下", r"以上修改", r"本版为\s*AI",
]
PLACEHOLDERS = [r"【待确认[^】]*】", r"【待补充[^】]*】", r"\bTBD\b", r"\bXX+\b"]
CAPTION_TABLE = re.compile(r"^表\s*(\d+|[A-Z])-(\d+)\s+\S")
CAPTION_FIGURE = re.compile(r"^图\s*(\d+|[A-Z])-(\d+)\s+\S")


class Report:
    def __init__(self):
        self.errors = []
        self.warnings = []
        self.notes = []

    def error(self, code, msg):
        self.errors.append({"code": code, "message": msg})

    def warn(self, code, msg):
        self.warnings.append({"code": code, "message": msg})

    def note(self, msg):
        self.notes.append(msg)


def _body_items(doc):
    items = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            items.append(("p", Paragraph(child, doc)))
        elif child.tag == qn("w:tbl"):
            items.append(("tbl", Table(child, doc)))
    return items


def _style_eastasia(doc, style_name):
    try:
        style = doc.styles[style_name]
    except KeyError:
        return None
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return None
    return rfonts.get(qn("w:eastAsia"))


def check_page(doc, rep):
    brief = len(doc.sections) == 1
    if len(doc.sections) not in (1, 3):
        rep.error("sections", f"分节数应为 3（完整版式）或 1（短文档版式），实际 {len(doc.sections)}")
        return brief
    for i, sec in enumerate(doc.sections):
        if abs(sec.page_width.cm - 21.0) > 0.05 or abs(sec.page_height.cm - 29.7) > 0.05:
            rep.error("page-size", f"第 {i + 1} 节不是 A4：{sec.page_width.cm:.1f}×{sec.page_height.cm:.1f}cm")
        for name, value, expect in (
            ("上边距", sec.top_margin.cm, 2.54),
            ("下边距", sec.bottom_margin.cm, 2.54),
            ("左边距", sec.left_margin.cm, 3.0),
            ("右边距", sec.right_margin.cm, 2.5),
        ):
            if abs(value - expect) > 0.06:
                rep.warn("page-margin", f"第 {i + 1} 节{name} {value:.2f}cm，规范为 {expect}cm")
    if brief:
        rep.note("短文档版式：单节、无封面与目录")
        return brief
    pn = doc.sections[1]._sectPr.find(qn("w:pgNumType"))
    if pn is None or pn.get(qn("w:fmt")) != "lowerRoman":
        rep.warn("page-number", "前置部分页码未设置为罗马数字")
    pn = doc.sections[2]._sectPr.find(qn("w:pgNumType"))
    if pn is None or pn.get(qn("w:fmt")) != "decimal":
        rep.warn("page-number", "正文页码未设置为阿拉伯数字并重新起编")
    return brief


def check_toc(doc, rep, brief=False):
    if brief:
        return
    xml = doc.element.body.xml
    if 'TOC \\o "1-3"' not in xml and "TOC \\o" not in xml:
        rep.error("toc", "缺少自动目录域（TOC）")
    settings = doc.settings.element
    upd = settings.find(qn("w:updateFields"))
    if upd is None or upd.get(qn("w:val")) not in ("true", "1"):
        rep.warn("toc", "未设置打开时自动更新域，目录可能为空")


def check_revisions(doc, rep):
    for table in doc.tables:
        header = " ".join(c.text.strip() for c in table.rows[0].cells)
        if "版本号" in header and "修订说明" in header:
            if len(table.rows) < 2:
                rep.warn("revision", "修订记录表只有表头，没有记录行")
            return
    rep.error("revision", "缺少“修订记录”表（版本号/日期/修订人/修订说明/审核人）")


def check_styles(doc, rep):
    used = {}
    for kind, obj in _body_items(doc):
        if kind != "p":
            continue
        if not obj.text.strip():
            continue
        name = obj.style.name
        used[name] = used.get(name, 0) + 1
        if name not in ALLOWED_STYLES:
            rep.error("style", f"使用了规范外样式“{name}”：{obj.text.strip()[:30]}")
    for style_name, font in FONT_STYLES.items():
        east = _style_eastasia(doc, style_name)
        if east is None:
            rep.error("font", f"样式“{style_name}”未设置中文字体（eastAsia）")
        elif east != font:
            rep.warn("font", f"样式“{style_name}”中文字体为 {east}，规范为 {font}")
    rep.note("段落样式分布：" + "，".join(f"{k}×{v}" for k, v in sorted(used.items())))


def check_headings(doc, rep):
    counters = [0, 0, 0]
    appendix = False
    chapter = "0"
    first = True
    for kind, obj in _body_items(doc):
        if kind != "p" or obj.style.name not in ("Heading 1", "Heading 2", "Heading 3"):
            continue
        text = obj.text.strip()
        level = int(obj.style.name[-1])
        if first and level != 1 and not re.match(r"^附录", text):
            rep.error("heading", f"文档第一个标题不是一级标题：{text[:40]}")
        first = False
        if level == 1:
            if re.match(r"^附录\s*[A-Za-z一二三四五六七八九十]+", text):
                appendix = True
                chapter = text.split()[1] if len(text.split()) > 1 else "A"
                counters = [0, 0, 0]
                continue
            appendix = False
            counters[0] += 1
            counters[1] = counters[2] = 0
            chapter = str(counters[0])
            expected = f"{counters[0]} "
        elif appendix:
            counters[level - 1] += 1
            prefix = chapter + "." + ".".join(str(c) for c in counters[1:level])
            expected = prefix + " "
        else:
            counters[level - 1] += 1
            expected = ".".join(str(c) for c in counters[:level]) + " "
        if not text.startswith(expected):
            rep.error("heading", f"标题编号不连续：应为“{expected.strip()} …”，实际“{text}”")


def check_captions(doc, rep):
    items = _body_items(doc)
    tbl_seen = {}
    fig_seen = {}
    for idx, (kind, obj) in enumerate(items):
        if kind != "p":
            continue
        text = obj.text.strip()
        style = obj.style.name
        if style == "PM表题":
            m = CAPTION_TABLE.match(text)
            if not m:
                rep.error("caption", f"表题格式错误：{text[:40]}")
                continue
            chapter, num = m.group(1), int(m.group(2))
            tbl_seen.setdefault(chapter, 0)
            if num != tbl_seen[chapter] + 1:
                rep.error("caption", f"表编号不连续：{text[:40]}")
            tbl_seen[chapter] = num
            if idx + 1 >= len(items) or items[idx + 1][0] != "tbl":
                rep.error("caption", f"表题下方没有紧邻的表格：{text[:40]}")
        elif style == "PM图题":
            m = CAPTION_FIGURE.match(text)
            if not m:
                rep.error("caption", f"图题格式错误：{text[:40]}")
                continue
            chapter, num = m.group(1), int(m.group(2))
            fig_seen.setdefault(chapter, 0)
            if num != fig_seen[chapter] + 1:
                rep.error("caption", f"图编号不连续：{text[:40]}")
            fig_seen[chapter] = num
    for table in doc.tables:
        grid = table._tbl.find(qn("w:tblGrid"))
        if grid is None:
            continue
        width = sum(int(c.get(qn("w:w")) or 0) for c in grid.findall(qn("w:gridCol")))
        if width and width > CONTENT_WIDTH_CM * 567 + 20:
            rep.warn("table-width", f"表格宽度 {width / 567:.1f}cm 超出正文宽度 {CONTENT_WIDTH_CM}cm")


def check_text(doc, rep):
    texts = []
    for kind, obj in _body_items(doc):
        if kind == "p":
            texts.append(obj.text)
        else:
            for row in obj.rows:
                texts.extend(cell.text for cell in row.cells)
    hits = []
    placeholders = []
    for text in texts:
        for pattern in BANNED:
            if re.search(pattern, text, re.I):
                hits.append(text.strip()[:60])
        for pattern in PLACEHOLDERS:
            placeholders.extend(re.findall(pattern, text))
    if hits:
        rep.error("process-note", "出现过程性说明或修订记录口吻：" + "；".join(hits[:5]))
    if placeholders:
        rep.warn("placeholder", f"仍有 {len(placeholders)} 处占位符，需在附录“待确认事项”中登记："
                 + "、".join(placeholders[:6]))


def check_header(doc, rep, brief=False):
    section = doc.sections[0] if brief else (doc.sections[1] if len(doc.sections) > 1 else None)
    if section is None:
        return
    header_text = " ".join(p.text for p in section.header.paragraphs)
    cover_version = ""
    for p in doc.paragraphs:
        m = re.search(r"版\s*本\s*号\s*[：:]\s*(\S+)", p.text)
        if m:
            cover_version = m.group(1)
            break
    hv = re.search(r"(V\d+\.\d+)", header_text)
    if cover_version and (not hv or hv.group(1) != cover_version):
        rep.error("version", f"页眉版本与封面版本不一致：页眉 {hv.group(1) if hv else '无'}，封面 {cover_version}")


def main() -> int:
    parser = argparse.ArgumentParser(description="检查项目管理过程文档 DOCX 的规范性")
    parser.add_argument("docx", help="待检查的 .docx")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    args = parser.parse_args()

    doc = Document(args.docx)
    rep = Report()
    brief = check_page(doc, rep)
    check_toc(doc, rep, brief)
    check_revisions(doc, rep)
    check_styles(doc, rep)
    check_headings(doc, rep)
    check_captions(doc, rep)
    check_text(doc, rep)
    check_header(doc, rep, brief)

    result = {
        "file": args.docx,
        "ok": not rep.errors,
        "errors": rep.errors,
        "warnings": rep.warnings,
        "notes": rep.notes,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for item in rep.errors:
            print(f"ERROR [{item['code']}] {item['message']}")
        for item in rep.warnings:
            print(f"WARN  [{item['code']}] {item['message']}")
        for line in rep.notes:
            print(f"NOTE  {line}")
        print(f"RESULT {'PASS' if not rep.errors else 'FAIL'} "
              f"errors={len(rep.errors)} warnings={len(rep.warnings)}")
    return 0 if not rep.errors else 1


if __name__ == "__main__":
    sys.exit(main())
