# soft-copyright-docs

[简体中文](README.md) · English

Generate the four documents required for a **Chinese software copyright registration** from a source tree, with one command.

```bash
python -m ruanzhu
```

One command, four Word documents:

| Document | Contents |
|------|------|
| **Source-code document** | 55 lines per page, automatically the first 30 + last 30 pages, header with software name, version and page number, continuous line numbers |
| **Software specification** | Cover + auto table of contents + overview / runtime environment / module descriptions / glossary |
| **User manual** | Cover + auto table of contents + install & launch / step-by-step instructions per module (with screenshot placeholders) / notes |
| **Screenshot document** | Cover + module index table + one screenshot per page (dashed placeholder frames) |

Chinese fonts, margins, line spacing and table shading are already set. Submit the output directly, or paste a few screenshots into the placeholders.

---

## Why this exists

The formatting rules for a software copyright filing are fiddly in a way that is **mandatory but completely mechanical**:

- The source-code document must contain the **first 30 + last 30 pages**, with a minimum number of lines per page and the software name/version in the header.
- The specification and the user manual are separate documents, each with a cover and table of contents.
- If the office asks for corrections, an additional screenshot document is required.

Doing this by hand takes an entire afternoon — and then you do it all over again for the next piece of software.

This tool **turns the rules into code**: describe your software in one JSON file and let it handle the layout.

---

## Quick start

Requires **Python 3.9+** and only `python-docx`.

```bash
git clone https://github.com/zyx0407/soft-copyright-docs.git
cd soft-copyright-docs
pip install python-docx
```

Then `python -m ruanzhu` works from inside the repository.

To use it from any directory, install the package itself (this registers a `ruanzhu` command):

```bash
pip install .
ruanzhu --help
```

### 1. Generate a config template

```bash
python -m ruanzhu --init
```

This creates `project.json` in the current directory.

### 2. Fill in the config

Open `project.json` and describe your software. The minimum that runs:

```json
{
  "name": "Your software name",
  "version": "V1.0",
  "summary": "One paragraph describing what the software does.",
  "modules": [
    {
      "id": "1",
      "name": "Module name",
      "desc": "What this module does.",
      "steps": ["First step.", "Second step."]
    }
  ],
  "source": {
    "dirs": ["src"],
    "include": ["**/*.py"]
  }
}
```

### 3. Generate

```bash
python -m ruanzhu
```

Output goes to `软著申请文档/` by default.

### Try the example first

The repository includes a runnable demo project:

```bash
python -m ruanzhu -c examples/demo/project.json
```

---

## Check before you submit

```bash
python verify_materials.py 软著申请文档
```

It checks the half a machine can check: all four files are present and not empty shells; the software name and version agree across all four; page numbers in the source-code document are continuous; every page title matches its page footer; line numbers run from 1 to the end with no gap.

```bash
python verify_materials.py 软著申请文档 --name "My Software" --version V1.0   # also cross-check name/version
```

`FAIL` means the material should not be submitted yet; `WARN` is worth a look. Table-of-contents fields and screenshots cannot be checked automatically and are listed at the end of the report.

---

## Use it with an AI assistant

The repository root ships a `SKILL.md` (Agent Skill format: `name`, `description`, workflow). Once it is installed into a compatible assistant, the assistant asks for the required fields, writes `project.json`, generates the documents and runs the self-check above:

```bash
# example: clone into an assistant's skills directory
git clone https://github.com/zyx0407/soft-copyright-docs.git ~/.claude/skills/soft-copyright-docs
```

Working by hand is exactly the same — `SKILL.md` is just the Q&A version of this README.

---

## Config fields

### Basics

| Field | Required | Description |
|------|:----:|------|
| `name` | ✅ | Software name, used on the cover, headers and in the body |
| `name_en` | | English name |
| `version` | | Version, default `V1.0` |
| `language` | | Development language, default `Python 3.10+`, printed on the source-code cover |
| `developer` | | Developer / organisation; the cover shows "(please fill in)" when empty |
| `completion_date` | | Completion date; same as above when empty |
| `output_dir` | | Output directory, default `软著申请文档` |
| `summary` | | One-paragraph description |
| `hardware` / `software` | | Hardware / software environment, arrays of strings |

### Content

| Field | Description |
|------|------|
| `features` | Core features, `[{"title": "name", "detail": "description"}]` |
| `modules` | **Functional modules**; all four documents are driven by this — see below |
| `glossary` | Glossary, `[{"term": "abbr", "full": "full name", "desc": "description"}]` |

`modules` is the most important field — it determines the manual chapters, every page of the screenshot document, and the index table:

```json
{
  "id": "1",
  "name": "Data import and preview",
  "desc": "Import experiment data files, merge them automatically, preview the first rows.",
  "steps": [
    "Click the \"Data import\" button on the left.",
    "Choose one or more data files and click \"Open\"."
  ]
}
```

### Source-code selection

| Field | Default | Description |
|------|------|------|
| `dirs` | `["."]` | Source directories, relative to the config file |
| `include` | `["**/*.py"]` | Files to collect |
| `exclude` | caches, virtualenvs, build output | Exclusion rules |
| `order` | `[]` | Files to place first (put entry points here) |
| `lines_per_page` | `55` | Lines per page in the source document |
| `first_pages` | `30` | Pages taken from the beginning |
| `last_pages` | `30` | Pages taken from the end |

> These defaults follow common filing practice: submit the first and last 30 consecutive pages of source code (60 pages total), at least 50 lines per page. **If the code is shorter than 60 pages, everything is submitted** and no error is raised.
> Check the latest guidance from the Copyright Protection Center of China; individual receiving offices may differ slightly.

### Command-line options

```bash
python -m ruanzhu --init              # generate a config template
python -m ruanzhu                     # generate everything from ./project.json
python -m ruanzhu -c my-config.json   # use a specific config file
python -m ruanzhu -o output-dir       # override the output directory
python -m ruanzhu --only source spec  # generate only the listed documents
```

Document names: `source`, `spec`, `manual`, `screenshots`.

---

## FAQ

**The generated table of contents is empty?**
Word's TOC is a field. Open the document, right-click the TOC, choose **Update Field → Update entire table**. The field codes are already inserted; Word just doesn't render them during headless generation.

**What if the code is shorter than 60 pages?**
Generate anyway — all pages are submitted, no blank pages are added.

**Fonts look wrong / Chinese renders in the wrong font?**
Chinese fonts must set `w:eastAsia` in addition to the Latin font, otherwise Word renders Chinese with a default font and you get a mismatched layout. This tool binds it everywhere fonts are set; keep that pattern if you modify the code.

**How do I add project-specific chapters?**
`steps` inside `modules` is a free-form string array — add as many as you like. For additional sections, edit `generate()` in `ruanzhu/manual_doc.py`.

---

## Project layout

```
soft-copyright-docs/
├── ruanzhu/
│   ├── __main__.py        CLI entry
│   ├── config.py          config model + filing defaults
│   ├── docx_helpers.py    Word layout primitives (pure layout, knows nothing about copyright)
│   ├── source_doc.py      source-code document
│   ├── spec_doc.py        software specification
│   ├── manual_doc.py      user manual
│   └── screenshot_doc.py  screenshot document
├── examples/demo/         runnable example project
├── tests/                 tests
├── SKILL.md               skill description for AI assistants
├── verify_materials.py    post-generation self-check
└── project.example.json   filled-in example config
```

The layering: `docx_helpers.py` only knows how to lay text out nicely; the four `*_doc.py` files only know which documents a filing needs and what goes in them. Changing the look touches one file; adding a new document touches one file.

Run the tests (no extra install):

```bash
python -m unittest discover -s tests -t .
```

---

## Scope

- This is a **layout tool**. It does not file the application, generate forms, or guarantee that the material passes review.
- Generated content comes **entirely from your `project.json`** — the tool never invents features. The more specific `summary`, `modules` and `steps` are, the better the material.
- Make sure you own the copyright of the source code you submit. For joint development, agree on ownership first.

## License

[MIT](LICENSE)
