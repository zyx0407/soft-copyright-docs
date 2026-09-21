"""命令行入口。

    python -m ruanzhu                     # 用当前目录的 project.json
    python -m ruanzhu -c 我的配置.json
    python -m ruanzhu --only source       # 只生成源程序文档
    python -m ruanzhu --init              # 生成一份配置模板
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__, manual_doc, screenshot_doc, source_doc, spec_doc
from .config import ConfigError, ProjectConfig

DOCUMENTS = {
    "source": ("源程序文档", source_doc.generate),
    "spec": ("软件说明书", spec_doc.generate),
    "manual": ("用户操作手册", manual_doc.generate),
    "screenshots": ("界面截图文档", screenshot_doc.generate),
}

DEFAULT_CONFIG_NAMES = ["project.json", "soft-copyright.json"]

TEMPLATE = """{
  "name": "示例软件名称",
  "name_en": "Example Software Name",
  "version": "V1.0",
  "language": "Python 3.10+",
  "developer": "",
  "completion_date": "",

  "output_dir": "软著申请文档",

  "summary": "用一段话说明这个软件是做什么的、解决什么问题。",

  "hardware": [
    "CPU：Intel Core i5 及以上",
    "内存：4 GB 及以上",
    "硬盘：可用空间 >= 500 MB"
  ],
  "software": [
    "操作系统：Windows 10 / 11（64 位）"
  ],

  "features": [
    {"title": "功能名称", "detail": "这个功能做什么。"}
  ],

  "glossary": [
    {"term": "缩写", "full": "全称", "desc": "说明。"}
  ],

  "modules": [
    {
      "id": "1",
      "name": "功能模块名称",
      "desc": "这个模块是干什么的。",
      "steps": [
        "第一步做什么。",
        "第二步做什么。"
      ]
    }
  ],

  "source": {
    "dirs": ["."],
    "include": ["**/*.py"],
    "exclude": ["**/__pycache__/**", "**/.venv/**", "**/build/**", "**/dist/**"],
    "order": [],
    "lines_per_page": 55,
    "first_pages": 30,
    "last_pages": 30
  }
}
"""


def find_config(explicit: str | None) -> Path:
    """定位配置文件。"""
    if explicit:
        path = Path(explicit)
        if not path.exists():
            raise SystemExit(f"[错误] 找不到配置文件：{path}")
        return path

    for name in DEFAULT_CONFIG_NAMES:
        candidate = Path.cwd() / name
        if candidate.exists():
            return candidate

    raise SystemExit(
        "[错误] 当前目录下没有找到 project.json。\n"
        "       可以用 --init 生成一份配置模板，或用 -c 指定配置文件路径。"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m ruanzhu",
        description="从源码目录一键生成软件著作权登记所需的四份材料。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python -m ruanzhu --init              生成配置模板\n"
            "  python -m ruanzhu                     用 project.json 生成全部材料\n"
            "  python -m ruanzhu --only source spec  只生成指定材料\n"
        ),
    )
    parser.add_argument("-c", "--config", metavar="文件",
                        help="配置文件路径（默认找 ./project.json）")
    parser.add_argument("-o", "--output", metavar="目录",
                        help="覆盖配置里的 output_dir")
    parser.add_argument("--only", nargs="+", choices=sorted(DOCUMENTS),
                        metavar="材料",
                        help="只生成指定材料："
                             + " / ".join(sorted(DOCUMENTS)))
    parser.add_argument("--init", action="store_true",
                        help="在当前目录生成一份配置模板后退出")
    parser.add_argument("--version", action="version",
                        version=f"soft-copyright-docs {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.init:
        target = Path.cwd() / "project.json"
        if target.exists():
            print(f"[跳过] {target} 已经存在，没有覆盖。")
            return 1
        target.write_text(TEMPLATE, encoding="utf-8")
        print(f"[OK] 配置模板已生成：{target}")
        print("     填好里面的字段，再运行 python -m ruanzhu 即可。")
        return 0

    try:
        cfg = ProjectConfig.load(find_config(args.config))
    except ConfigError as exc:
        print(f"[错误] {exc}")
        return 2

    if args.output:
        cfg.output_dir = args.output

    wanted = args.only or list(DOCUMENTS)

    print("=" * 60)
    print(f"  {cfg.name}  {cfg.version}")
    print(f"  输出目录：{cfg.output_path()}")
    print(f"  生成材料：{'、'.join(DOCUMENTS[k][0] for k in wanted)}")
    print("=" * 60)

    produced: list[Path] = []
    for key in wanted:
        label, func = DOCUMENTS[key]
        try:
            produced.append(func(cfg))
        except SystemExit as exc:
            print(str(exc))
            return 3

    print("\n" + "=" * 60)
    print(f"  完成，共生成 {len(produced)} 份材料：")
    for path in produced:
        size_kb = path.stat().st_size / 1024
        print(f"    {path.name}  ({size_kb:.0f} KB)")
    print("=" * 60)
    print("\n提示：打开 Word 文档后，在目录处右键 → 更新域 → 更新整个目录，")
    print("      即可生成带页码的完整目录。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
