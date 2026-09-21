"""配置模型，以及软著申报的格式默认值。

软著申报对材料格式有硬性要求（源代码文档前 30 页 + 后 30 页、每页 55 行等）。
本模块把这些要求固化成默认值，使用者只需要在 ``project.json`` 里描述自己的软件，
不必关心页数行数怎么排。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

# ──────────────────────────────────────────────
# 软著申报的格式要求（默认值，可按需覆盖）
# ──────────────────────────────────────────────

LINES_PER_PAGE = 55        # 源代码文档每页行数
FIRST_PART_PAGES = 30      # 取前 N 页
LAST_PART_PAGES = 30       # 取后 N 页（合计 60 页）

A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7

DEFAULT_EXCLUDES = [
    "**/__pycache__/**",
    "**/.venv/**",
    "**/venv/**",
    "**/build/**",
    "**/dist/**",
    "**/.git/**",
    "**/node_modules/**",
]

DEFAULT_INCLUDES = ["**/*.py"]


class ConfigError(Exception):
    """配置文件有问题时抛出。"""


@dataclass
class SourceConfig:
    """源代码文档的取材范围与分页参数。"""

    dirs: list[str] = field(default_factory=lambda: ["."])
    include: list[str] = field(default_factory=lambda: list(DEFAULT_INCLUDES))
    exclude: list[str] = field(default_factory=lambda: list(DEFAULT_EXCLUDES))
    order: list[str] = field(default_factory=list)
    lines_per_page: int = LINES_PER_PAGE
    first_pages: int = FIRST_PART_PAGES
    last_pages: int = LAST_PART_PAGES


@dataclass
class Module:
    """一个功能模块，用于生成截图索引表和操作手册的章节。"""

    id: str = ""
    name: str = ""
    desc: str = ""
    steps: list[str] = field(default_factory=list)


@dataclass
class Feature:
    """一条核心功能，用于软件说明书的「核心功能」小节。"""

    title: str = ""
    detail: str = ""


@dataclass
class Glossary:
    """一条术语，用于软件说明书的术语表。"""

    term: str = ""
    full: str = ""
    desc: str = ""


@dataclass
class ProjectConfig:
    """一个待申报软件的完整描述。"""

    name: str = "未命名软件"
    name_en: str = ""
    version: str = "V1.0"
    language: str = "Python 3.10+"
    developer: str = ""
    completion_date: str = ""
    output_dir: str = "软著申请文档"
    summary: str = ""
    features: list[Feature] = field(default_factory=list)
    modules: list[Module] = field(default_factory=list)
    glossary: list[Glossary] = field(default_factory=list)
    hardware: list[str] = field(default_factory=list)
    software: list[str] = field(default_factory=list)
    source: SourceConfig = field(default_factory=SourceConfig)

    # 运行期信息，不从 JSON 读
    root: Path = field(default_factory=Path.cwd)

    # ── 读取 ──────────────────────────────────

    @classmethod
    def load(cls, path: str | Path) -> "ProjectConfig":
        """从 JSON 文件读取配置。相对路径一律相对于配置文件所在目录。"""
        path = Path(path).resolve()
        if not path.exists():
            raise ConfigError(f"找不到配置文件：{path}")

        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise ConfigError(f"配置文件不是合法的 JSON：{exc}") from exc

        if not isinstance(raw, dict):
            raise ConfigError("配置文件的顶层必须是一个对象（{...}）")

        cfg = cls._from_dict(raw)
        cfg.root = path.parent
        return cfg

    @classmethod
    def _from_dict(cls, raw: dict) -> "ProjectConfig":
        known = {
            "name", "name_en", "version", "language", "developer",
            "completion_date", "output_dir", "summary",
            "features", "modules", "glossary", "hardware", "software",
            "source",
        }
        unknown = set(raw) - known
        if unknown:
            raise ConfigError(
                "配置里有不认识的字段："
                + "、".join(sorted(unknown))
                + "\n可用字段：" + "、".join(sorted(known))
            )

        src_raw = raw.get("source") or {}
        if not isinstance(src_raw, dict):
            raise ConfigError("source 必须是一个对象（{...}）")

        source = SourceConfig(
            dirs=list(src_raw.get("dirs", ["."])),
            include=list(src_raw.get("include", DEFAULT_INCLUDES)),
            exclude=list(src_raw.get("exclude", DEFAULT_EXCLUDES)),
            order=list(src_raw.get("order", [])),
            lines_per_page=int(src_raw.get("lines_per_page", LINES_PER_PAGE)),
            first_pages=int(src_raw.get("first_pages", FIRST_PART_PAGES)),
            last_pages=int(src_raw.get("last_pages", LAST_PART_PAGES)),
        )

        def _mods() -> list[Module]:
            out = []
            for i, m in enumerate(raw.get("modules") or [], start=1):
                if not isinstance(m, dict):
                    raise ConfigError("modules 里的每一项都必须是对象")
                out.append(Module(
                    id=str(m.get("id") or i),
                    name=m.get("name", ""),
                    desc=m.get("desc", ""),
                    steps=list(m.get("steps") or []),
                ))
            return out

        def _feats() -> list[Feature]:
            out = []
            for f in raw.get("features") or []:
                if isinstance(f, str):
                    out.append(Feature(title=f, detail=""))
                elif isinstance(f, dict):
                    out.append(Feature(
                        title=f.get("title", ""),
                        detail=f.get("detail", ""),
                    ))
                else:
                    raise ConfigError("features 里只能是字符串或对象")
            return out

        def _gloss() -> list[Glossary]:
            out = []
            for g in raw.get("glossary") or []:
                if not isinstance(g, dict):
                    raise ConfigError("glossary 里的每一项都必须是对象")
                out.append(Glossary(
                    term=g.get("term", ""),
                    full=g.get("full", ""),
                    desc=g.get("desc", ""),
                ))
            return out

        return cls(
            name=raw.get("name", "未命名软件"),
            name_en=raw.get("name_en", ""),
            version=raw.get("version", "V1.0"),
            language=raw.get("language", "Python 3.10+"),
            developer=raw.get("developer", ""),
            completion_date=raw.get("completion_date", ""),
            output_dir=raw.get("output_dir", "软著申请文档"),
            summary=raw.get("summary", ""),
            features=_feats(),
            modules=_mods(),
            glossary=_gloss(),
            hardware=list(raw.get("hardware") or []),
            software=list(raw.get("software") or []),
            source=source,
        )

    # ── 派生 ──────────────────────────────────

    def output_path(self) -> Path:
        """输出目录的绝对路径（相对配置文件所在目录）。"""
        out = Path(self.output_dir)
        return out if out.is_absolute() else (self.root / out)

    def source_files(self) -> list[Path]:
        """按 include / exclude / order 规则收集源文件。

        order 里列出的文件排在最前（按 order 的顺序），其余按路径排序。
        """
        found: dict[str, Path] = {}

        for d in self.source.dirs:
            base = Path(d)
            if not base.is_absolute():
                base = self.root / base
            if not base.exists():
                print(f"  [WARN] 源码目录不存在，跳过：{base}")
                continue

            for pattern in self.source.include:
                for p in base.glob(pattern):
                    if not p.is_file():
                        continue
                    rel = p.relative_to(self.root).as_posix()
                    if any(p.match(ex) or Path(rel).match(ex) for ex in self.source.exclude):
                        continue
                    found[rel] = p

        ordered: list[Path] = []
        for name in self.source.order:
            key = Path(name).as_posix()
            if key in found:
                ordered.append(found.pop(key))
            else:
                print(f"  [WARN] order 里列出的文件没找到：{name}")

        ordered.extend(found[k] for k in sorted(found))
        return ordered
