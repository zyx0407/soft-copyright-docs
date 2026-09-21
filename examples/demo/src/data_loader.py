"""数据读取与合并。

支持 .csv / .xlsx / .xls 三种格式，按学号对齐后合并成一张表。
"""

from __future__ import annotations

import csv
from pathlib import Path

SUPPORTED = (".csv", ".xlsx", ".xls")


def list_supported(folder: str | Path) -> list[Path]:
    """列出目录下所有受支持的数据文件。"""
    folder = Path(folder)
    if not folder.is_dir():
        raise NotADirectoryError(f"不是有效的目录：{folder}")
    return sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED
    )


def read_csv(path: str | Path, encoding: str = "utf-8-sig") -> list[dict]:
    """读取一个 CSV 文件，返回以表头为键的字典列表。

    使用 utf-8-sig 是为了吃掉 Excel 导出的 BOM，否则第一列列名会带一个
    看不见的 \\ufeff 字符，后面按列名取值全部取不到。
    """
    path = Path(path)
    try:
        with path.open("r", encoding=encoding, newline="") as f:
            reader = csv.DictReader(f)
            rows = [dict(row) for row in reader]
    except UnicodeDecodeError:
        with path.open("r", encoding="gbk", newline="") as f:
            reader = csv.DictReader(f)
            rows = [dict(row) for row in reader]

    if not rows:
        raise ValueError(f"{path.name} 里没有任何数据行")

    return rows


def read_xlsx(path: str | Path) -> list[dict]:
    """读取一个 Excel 文件。需要安装 openpyxl。"""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # noqa: BLE001
        raise RuntimeError(
            "读取 Excel 需要 openpyxl，请先执行：pip install openpyxl"
        ) from exc

    wb = load_workbook(Path(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise ValueError(f"{Path(path).name} 是空表")

    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    out = []
    for row in rows[1:]:
        if all(c is None for c in row):
            continue
        out.append({header[i]: row[i] for i in range(min(len(header), len(row)))})
    return out


def read_any(path: str | Path) -> list[dict]:
    """按扩展名分派到对应的读取函数。"""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return read_csv(path)
    if suffix in (".xlsx", ".xls"):
        return read_xlsx(path)
    raise ValueError(f"不支持的文件格式：{suffix}")


def merge_by_key(tables: list[list[dict]], key: str) -> tuple[list[str], list[dict]]:
    """把多张表按 key 列合并成一张。

    列名取所有表的并集（保持首次出现的顺序），缺少的格子留空字符串。
    """
    columns: list[str] = []
    for table in tables:
        for row in table:
            for name in row:
                if name not in columns:
                    columns.append(name)

    if key not in columns:
        raise KeyError(
            f"合并用的键列“{key}”不在任何一张表里。"
            f"实际读到的列有：{'、'.join(columns)}"
        )

    merged: dict[str, dict] = {}
    for table in tables:
        for row in table:
            k = str(row.get(key, "")).strip()
            if not k:
                continue
            bucket = merged.setdefault(k, {c: "" for c in columns})
            bucket[key] = k
            for name, value in row.items():
                if value not in (None, ""):
                    bucket[name] = value

    ordered = [merged[k] for k in sorted(merged, key=_natural_key)]
    return columns, ordered


def _natural_key(value: str):
    """让 2 排在 10 前面，而不是按字符串比较。"""
    return tuple(
        int(part) if part.isdigit() else part
        for part in _split_digits(value)
    )


def _split_digits(text: str):
    import re
    return [p for p in re.split(r"(\d+)", text) if p != ""]
