"""学生成绩统计分析工具 —— 命令行入口。

用法：
    python main.py summary 数据目录 [--key 学号]
    python main.py export  数据目录 输出.docx [--key 学号]
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import data_loader
import report
import stats


def cmd_summary(args) -> int:
    """读取目录下所有数据文件，合并后输出描述性统计。"""
    tables = []
    for path in data_loader.list_supported(args.folder):
        print(f"  读入 {path.name} ...")
        try:
            tables.append(data_loader.read_any(path))
        except Exception as exc:  # noqa: BLE001
            print(f"  [跳过] {path.name}：{exc}")

    if not tables:
        print("没有读到任何数据。")
        return 1

    columns, rows = data_loader.merge_by_key(tables, args.key)
    print(f"\n合并完成：{len(rows)} 条记录，{len(columns)} 个字段\n")

    numeric_columns = _pick_numeric_columns(rows, columns, args.key)
    if not numeric_columns:
        print("没有找到可分析的数值列。")
        return 1

    summaries = [
        stats.summarize(col, [row.get(col) for row in rows])
        for col in numeric_columns
    ]

    header = report.HEADERS
    print("  ".join(f"{h:>10}" for h in header))
    for line in report.summary_to_rows(summaries):
        print("  ".join(f"{v:>10}" for v in line))

    print()
    for line in report.describe(summaries):
        print("  " + line)

    for s in summaries:
        found = stats.find_outliers(
            [row.get(s.name) for row in rows], method=args.method)
        if found:
            print(f"  [异常] {s.name} 有 {len(found)} 个异常值：{found}")

    return 0


def cmd_export(args) -> int:
    """分析并把结果导出为 Word 报告。"""
    tables = [data_loader.read_any(p)
              for p in data_loader.list_supported(args.folder)]
    if not tables:
        print("没有读到任何数据。")
        return 1

    columns, rows = data_loader.merge_by_key(tables, args.key)
    numeric_columns = _pick_numeric_columns(rows, columns, args.key)
    summaries = [
        stats.summarize(col, [row.get(col) for row in rows])
        for col in numeric_columns
    ]

    table_lines = ["    ".join(report.HEADERS)]
    for line in report.summary_to_rows(summaries):
        table_lines.append("    ".join(line))

    sections = {
        "一、数据概览": [
            f"本次分析共合并 {len(tables)} 个数据文件，"
            f"得到 {len(rows)} 条记录，涉及 {len(columns)} 个字段。",
            f"其中参与统计分析的数值字段有 {len(numeric_columns)} 个。",
        ],
        "二、描述性统计": table_lines,
        "三、结论": report.describe(summaries),
    }

    out = report.build_report(args.title, sections, args.output)
    print(f"[OK] 报告已生成：{out}")
    return 0


def _pick_numeric_columns(rows: list[dict], columns: list[str],
                          key: str) -> list[str]:
    """挑出真正的数值列：至少一半的值能转成数字。"""
    picked = []
    for col in columns:
        if col == key:
            continue
        values = [row.get(col) for row in rows]
        numbers = stats.to_numbers(values)
        if numbers and len(numbers) >= max(1, len(rows) // 2):
            picked.append(col)
    return picked


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="学生成绩统计分析工具",
    )
    parser.add_argument("--key", default="学号",
                        help="合并不同文件时使用的键列（默认：学号）")
    parser.add_argument("--method", default="iqr", choices=["iqr", "sigma"],
                        help="异常值判据（默认：iqr）")

    sub = parser.add_subparsers(dest="command", required=True)

    p1 = sub.add_parser("summary", help="输出描述性统计")
    p1.add_argument("folder", help="数据文件所在目录")
    p1.set_defaults(func=cmd_summary)

    p2 = sub.add_parser("export", help="导出 Word 报告")
    p2.add_argument("folder", help="数据文件所在目录")
    p2.add_argument("output", help="输出的 .docx 路径")
    p2.add_argument("--title", default="学生成绩统计分析报告",
                    help="报告标题")
    p2.set_defaults(func=cmd_export)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
