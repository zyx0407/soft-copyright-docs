# soft-copyright-docs

[English](README.en.md) · 简体中文

从源码目录一键生成**软件著作权登记**所需的四份材料。

```bash
python -m ruanzhu
```

一条命令，得到四份 Word 文档：

| 材料 | 内容 |
|------|------|
| **源程序文档** | 每页 55 行，自动取前 30 页 + 后 30 页，页眉带软件名、版本号、页码，每行连续行号 |
| **软件说明书** | 封面 + 自动目录 + 软件概述 / 运行环境 / 功能模块说明 / 术语表 |
| **用户操作手册** | 封面 + 自动目录 + 安装启动 / 逐模块操作步骤（带截图占位框）/ 注意事项 |
| **界面截图文档** | 封面 + 功能模块索引表 + 一页一张截图（虚线框占位） |

中文字体、页边距、行距、表格底纹都排好了，生成之后直接提交，或者往占位框里贴几张截图就行。

---

## 为什么写这个

软著申报的材料格式要求很碎，而且是「**必须对、但完全没有技术含量**」的那种碎：

- 源程序文档要**前 30 页 + 后 30 页**，每页行数有下限，页眉要写明软件名和版本号
- 说明书和操作手册要分开写，都带封面和目录
- 被要求补正的时候，还得额外提交一份界面截图文档

这些事手动做一次要花掉一整个下午，而且下次申请另一个软件还要再来一遍。

这个工具就是把这些规则**固化成代码**：你只需要用一份 JSON 描述自己的软件，剩下的排版交给它。

---

## 快速开始

需要 **Python 3.9+**，只依赖 `python-docx`。

```bash
git clone https://github.com/zyx0407/soft-copyright-docs.git
cd soft-copyright-docs
pip install python-docx
```

这样在**仓库目录里**就能直接用 `python -m ruanzhu`。

如果你想在任何目录下都能用，再装一下它自己（会注册一个 `ruanzhu` 命令）：

```bash
pip install .
ruanzhu --help
```

### 1. 生成一份配置模板

```bash
python -m ruanzhu --init
```

会在当前目录生成 `project.json`。

### 2. 填写配置

打开 `project.json`，把你的软件描述进去。最少只要填这几项就能跑：

```json
{
  "name": "你的软件名称",
  "version": "V1.0",
  "summary": "用一段话说明这个软件是做什么的。",
  "modules": [
    {
      "id": "1",
      "name": "功能模块名称",
      "desc": "这个模块是干什么的。",
      "steps": ["第一步做什么。", "第二步做什么。"]
    }
  ],
  "source": {
    "dirs": ["src"],
    "include": ["**/*.py"]
  }
}
```

### 3. 生成

```bash
python -m ruanzhu
```

输出默认落在 `软著申请文档/` 目录下。

### 先试试示例

仓库里带了一个可以直接跑通的示例项目：

```bash
python -m ruanzhu -c examples/demo/project.json
```

---

## 生成完先自检

```bash
python verify_materials.py 软著申请文档
```

查机器能查的那一半：四份文件在不在、是不是空壳、四份的软件名与版本号是否一致、源程序文档页码是否连续、每页标题与页脚对不对得上、代码行号是否从 1 连续编到末尾。

```bash
python verify_materials.py 软著申请文档 --name 我的软件 --version V1.0   # 顺带核名称/版本
```

有 `FAIL` 说明材料不能交，`WARN` 看一眼再决定。目录域和截图它查不了，退出时会单独提示。

---

## 交给 AI 助手用

仓库根带一份 `SKILL.md`（Agent Skill 格式：`name` + `description` + 工作流）。给支持该格式的助手装上之后，它会按流程问齐字段、写 `project.json`、生成，再跑上面的自检：

```bash
# 例：克隆到助手的 skills 目录
git clone https://github.com/zyx0407/soft-copyright-docs.git ~/.claude/skills/soft-copyright-docs
```

手动用也完全一样，`SKILL.md` 就是这份 README 的「问答版」。

---

## 配置字段

### 基本信息

| 字段 | 必填 | 说明 |
|------|:----:|------|
| `name` | ✅ | 软件名称，会出现在封面、页眉和文档正文里 |
| `name_en` | | 英文名称 |
| `version` | | 版本号，默认 `V1.0` |
| `language` | | 开发语言，默认 `Python 3.10+`，会写进源程序文档封面 |
| `developer` | | 开发单位。留空时封面显示「（请填写）」 |
| `completion_date` | | 完成日期。留空时同上 |
| `output_dir` | | 输出目录，默认 `软著申请文档` |
| `summary` | | 软件简介，一段话 |
| `hardware` / `software` | | 硬件环境 / 软件环境，字符串数组，每项一条 |

### 内容

| 字段 | 说明 |
|------|------|
| `features` | 核心功能，`[{"title": "功能名", "detail": "说明"}]` |
| `modules` | **功能模块**，四份材料都靠它驱动，详见下 |
| `glossary` | 术语表，`[{"term": "缩写", "full": "全称", "desc": "说明"}]` |

`modules` 是最重要的字段 —— 它决定操作手册的章节、截图文档的每一页、以及索引表：

```json
{
  "id": "1",
  "name": "数据导入与预览",
  "desc": "导入实验数据文件并自动合并，预览前若干行。",
  "steps": [
    "点击左侧的“数据导入”按钮。",
    "选择一个或多个数据文件，点击“打开”。"
  ]
}
```

### 源程序取材

| 字段 | 默认 | 说明 |
|------|------|------|
| `dirs` | `["."]` | 源码所在目录，相对配置文件位置 |
| `include` | `["**/*.py"]` | 收集哪些文件 |
| `exclude` | 排除缓存、虚拟环境、构建产物 | 排除规则 |
| `order` | `[]` | 指定排在最前面的文件（建议放入口文件） |
| `lines_per_page` | `55` | 源代码文档每页行数 |
| `first_pages` | `30` | 取前多少页 |
| `last_pages` | `30` | 取后多少页 |

> 这三个默认值是照软著申报的通行做法设的：源程序文档提交前、后各连续 30 页，共 60 页，每页不少于 50 行。**代码总页数不足 60 页时会全部提交**，不会报错。
> 具体要求以中国版权保护中心的最新说明为准，各受理点口径可能略有差异。

### 命令参数

```bash
python -m ruanzhu --init              # 生成配置模板
python -m ruanzhu                     # 用 ./project.json 生成全部材料
python -m ruanzhu -c 我的配置.json     # 指定配置文件
python -m ruanzhu -o 输出目录          # 覆盖输出目录
python -m ruanzhu --only source spec  # 只生成指定材料
```

可选材料名：`source`（源程序文档）、`spec`（软件说明书）、`manual`（用户操作手册）、`screenshots`（界面截图文档）。

---

## 常见问题

**生成的目录是空的？**
Word 的目录是「域」，需要打开文档后手动更新一次：在目录处右键 → **更新域** → **更新整个目录**。文档里已经插好了域代码，只是 Word 不会在无头生成时预先渲染。

**代码不够 60 页怎么办？**
直接生成就行，页数不够时会全部提交，不会补空页。

**字体不对 / 中文变成别的字体了？**
中文字体必须同时绑定 `w:eastAsia` 属性，否则 Word 会用默认字体渲染中文，出现「英文宋体 + 中文宋体以外的字体」的乱版。这个工具在每一处设字体的地方都绑定了，如果你自己改代码，记得保持这个写法。

**想加自己项目特有的章节？**
`modules` 里的 `steps` 是自由的字符串数组，加多少条都行。如果还需要别的章节，改 `ruanzhu/manual_doc.py` 里的 `generate()` 即可。

---

## 项目结构

```
soft-copyright-docs/
├── ruanzhu/
│   ├── __main__.py        命令行入口
│   ├── config.py          配置模型 + 申报格式默认值
│   ├── docx_helpers.py    Word 排版原语（纯排版，不认识「软著」）
│   ├── source_doc.py      源程序文档
│   ├── spec_doc.py        软件说明书
│   ├── manual_doc.py      用户操作手册
│   └── screenshot_doc.py  界面截图文档
├── examples/demo/         可直接跑通的示例项目
├── tests/                 测试
├── SKILL.md               给 AI 助手用的 skill 说明
├── verify_materials.py    生成后的机器自检
└── project.example.json   填好的配置示例
```

分层的用意：`docx_helpers.py` 只负责「怎么把字排好看」，四个 `*_doc.py` 只负责「软著要哪几份材料、每份写什么」。要换排版风格只改一个文件，要加一种材料只加一个文件。

运行测试（不需要额外安装任何东西）：

```bash
python -m unittest discover -s tests -t .
```

---

## 说明与边界

- 这是一个**排版工具**，不代办申请、不生成申请表，也不保证材料一定通过审查。
- 生成的内容**完全来自你的 `project.json`**，工具不会自己编造软件功能。所以 `summary`、`modules`、`steps` 写得越具体，材料质量越高。
- 请确保你对自己提交的源代码拥有著作权。如果是合作开发，先确认权利归属再提交。

## License

[MIT](LICENSE)
