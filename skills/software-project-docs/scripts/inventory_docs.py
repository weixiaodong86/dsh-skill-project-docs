# -*- coding: utf-8 -*-
"""盘点一个目录里已有的项目文档，为重构建模提供输入清单。

用法：
    python -B scripts/inventory_docs.py <目录> [-o inventory.json] [--max-files 200]

输出：stdout 一张概览表；-o 指定时另存 JSON（供代理读取完整标题结构）。
支持 .docx / .md / .markdown / .txt；PDF 与 .doc 只登记、提示先转换。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime

SUPPORTED = {".docx", ".md", ".markdown", ".txt"}
CONVERT_FIRST = {".pdf", ".doc", ".wps", ".rtf"}
SKIP_DIRS = {"node_modules", ".git", "__pycache__", ".venv", "venv", "_source"}


def _docx_info(path):
    from docx import Document

    doc = Document(path)
    headings = []
    chars = 0
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        chars += len(text)
        style = paragraph.style.name or ""
        match = re.search(r"(\d+)$", style) if style.lower().startswith("heading") else None
        if match:
            headings.append({"level": int(match.group(1)), "text": text[:120]})
        elif style.startswith("标题"):
            num = re.search(r"(\d+)$", style)
            if num:
                headings.append({"level": int(num.group(1)), "text": text[:120]})
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                chars += len(cell.text.strip())
    return {"chars": chars, "tables": len(doc.tables), "headings": headings}


def _text_info(path):
    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        content = handle.read()
    headings = []
    tables = 0
    in_table = False
    for line in content.splitlines():
        match = re.match(r"^(#{1,4})\s+(.+)$", line)
        if match:
            headings.append({"level": len(match.group(1)), "text": match.group(2).strip()[:120]})
        is_row = line.strip().startswith("|")
        if is_row and not in_table:
            tables += 1
        in_table = is_row
    return {"chars": len(content), "tables": tables, "headings": headings}


def _info(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return _docx_info(path)
    return _text_info(path)


def scan(root, max_files=200):
    documents = []
    unsupported = []
    for current, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith(".")]
        for name in sorted(files):
            if name.startswith("~$"):
                continue
            ext = os.path.splitext(name)[1].lower()
            full = os.path.join(current, name)
            rel = os.path.relpath(full, root)
            if ext in CONVERT_FIRST:
                unsupported.append({"path": rel, "reason": f"{ext} 需先转换为 docx 或文本"})
                continue
            if ext not in SUPPORTED:
                continue
            if len(documents) >= max_files:
                return documents, unsupported, True
            record = {
                "path": rel,
                "name": name,
                "ext": ext,
                "bytes": os.path.getsize(full),
                "mtime": datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d"),
                "readable": True,
                "error": None,
                "chars": 0,
                "tables": 0,
                "headings": [],
            }
            try:
                record.update(_info(full))
            except Exception as error:  # 损坏文件只登记，不中断盘点
                record["readable"] = False
                record["error"] = f"{type(error).__name__}: {error}"
            documents.append(record)
    return documents, unsupported, False


def main():
    parser = argparse.ArgumentParser(description="盘点目录中的项目文档")
    parser.add_argument("root", help="待盘点的目录")
    parser.add_argument("-o", "--out", help="输出 JSON 文件")
    parser.add_argument("--max-files", type=int, default=200)
    args = parser.parse_args()

    if not os.path.isdir(args.root):
        print(json.dumps({"ok": False, "error": f"目录不存在：{args.root}"}, ensure_ascii=False))
        return 2

    documents, unsupported, truncated = scan(os.path.abspath(args.root), args.max_files)
    result = {
        "ok": True,
        "root": os.path.abspath(args.root),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count": len(documents),
        "truncated": truncated,
        "documents": documents,
        "unsupported": unsupported,
    }

    print(f"{'文件':<44}{'字数':>8}{'表':>5}{'标题':>6}  {'更新':<12}状态")
    for item in documents:
        status = "可读" if item["readable"] else "损坏"
        name = item["path"]
        if len(name) > 42:
            name = "…" + name[-41:]
        print(f"{name:<44}{item['chars']:>8}{item['tables']:>5}{len(item['headings']):>6}  "
              f"{item['mtime']:<12}{status}")
    for item in unsupported:
        print(f"{item['path']:<44}{'':>8}{'':>5}{'':>6}  {'':<12}{item['reason']}")
    print(f"共 {len(documents)} 份可读文档" + ("（已截断）" if truncated else "")
          + (f"，{len(unsupported)} 份需先转换" if unsupported else ""))

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2)
        print(f"JSON 已写入：{args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
