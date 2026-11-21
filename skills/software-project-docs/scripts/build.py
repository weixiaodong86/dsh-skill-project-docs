# -*- coding: utf-8 -*-
"""构建入口：把受控 Markdown 内容源渲染为规范 DOCX。

用法：
    python -B scripts/build.py <content.md> -o <out.docx> [--version V1.1] [--date 2026-11-21]

脚本会把 assets/docx_kit.py 所在目录加入 sys.path，因此技能资源可以只读。
输出 JSON 摘要，供代理在回复中引用。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "assets"))

from docx_kit import build_document  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="生成软件研发项目管理过程文档 DOCX")
    parser.add_argument("content", help="内容源 .md 路径")
    parser.add_argument("-o", "--out", required=True, help="输出 .docx 路径")
    parser.add_argument("--version", help="覆盖 front matter 中的版本号")
    parser.add_argument("--date", help="覆盖 front matter 中的日期")
    parser.add_argument("--title", help="覆盖文档标题")
    parser.add_argument("--doc-code", help="覆盖文档编号")
    args = parser.parse_args()

    if not os.path.isfile(args.content):
        print(json.dumps({"ok": False, "error": f"内容源不存在: {args.content}"},
                         ensure_ascii=False))
        return 2

    overrides = {
        "version": args.version,
        "date": args.date,
        "title": args.title,
        "doc_code": args.doc_code,
    }
    info = build_document(args.content, args.out, overrides)
    info["ok"] = True
    info["bytes"] = os.path.getsize(args.out)
    print(json.dumps(info, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
