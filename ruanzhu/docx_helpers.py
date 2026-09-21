"""Word 排版工具函数。

这一层是纯排版的，不认识「软著」这个概念。四个文档生成器都建在它上面。

关于中文字体：python-docx 里只设 ``run.font.name`` 是不够的，
Word 会用默认字体渲染中文，出现「英文宋体 + 中文默认体」的乱版。
所以每一处设字体的地方都必须同时绑定 ``w:eastAsia``。
"""

from __future__ import annotations

import re

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor

# ── 常用字体 ──────────────────────────────────

SONG = "宋体"        # 正文
HEI = "黑体"         # 标题
MONO = "Consolas"    # 代码

GREY = RGBColor(0xAA, 0xAA, 0xAA)
DARK = RGBColor(0x33, 0x33, 0x33)

_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.S)


# ──────────────────────────────────────────────
# 底层：run 与单元格
# ──────────────────────────────────────────────

def add_run(paragraph, text, font_name=SONG, font_size=Pt(11),
            bold=False, color=None, italic=False):
    """往段落里加一个 run，并把中文字体一并绑定好。"""
    run = paragraph.add_run(text)
    run.font.name = font_name
    # 必须先设 name，rPr/rFonts 才会被创建出来
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    return run


def add_rich_text(paragraph, text, font_name=SONG, font_size=Pt(11),
                  color=None, bold=False):
    """写入一段可能含 ``**粗体**`` 标记的文字。

    标记会被解析成真正的粗体 run，而不是把星号原样打到 Word 里。
    """
    pos = 0
    matched = False
    for m in _BOLD_RE.finditer(text):
        matched = True
        if m.start() > pos:
            add_run(paragraph, text[pos:m.start()], font_name, font_size, bold, color)
        add_run(paragraph, m.group(1), font_name, font_size, True, color)
        pos = m.end()
    if pos < len(text) or not matched:
        add_run(paragraph, text[pos:], font_name, font_size, bold, color)


def set_cell_shading(cell, color_hex: str):
    """设置单元格底色。color_hex 形如 ``'D2691E'``。"""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_text(cell, text, font_name=SONG, font_size=Pt(10), bold=False,
                  color=None, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """把单元格内容整体替换成一段文字。"""
    cell.paragraphs[0].clear()
    p = cell.paragraphs[0]
    p.alignment = alignment
    add_rich_text(p, text, font_name, font_size, color, bold)
    return p


def set_col_widths(table, widths_cm: list[float]):
    """按厘米给表格各列设定宽度。"""
    for row in table.rows:
        for cell, width in zip(row.cells, widths_cm):
            cell.width = Cm(width)


# ──────────────────────────────────────────────
# 段落级
# ──────────────────────────────────────────────

def add_paragraph(doc, text="", font_name=SONG, font_size=Pt(11), bold=False,
                  alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6),
                  space_before=Pt(0), first_line_indent=None, color=None,
                  line_spacing=Pt(22)):
    """添加一个段落，支持 ``**粗体**`` 行内标记。"""
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    pf.space_after = space_after
    pf.space_before = space_before
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    add_rich_text(p, text, font_name, font_size, color, bold)
    return p


def add_indented_paragraph(doc, text, **kwargs):
    """正文段：首行缩进两个字符。"""
    kwargs.setdefault("first_line_indent", Cm(0.74))
    return add_paragraph(doc, text, **kwargs)


def add_heading_styled(doc, text, level=1):
    """标题。用内置 heading 样式，这样 Word 才能自动生成目录。"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = HEI
        run._element.rPr.rFonts.set(qn("w:eastAsia"), HEI)
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
    return h


def add_bullet(doc, text, font_size=Pt(11), indent_level=0):
    """项目符号列表项。"""
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    add_rich_text(p, text, SONG, font_size)
    p.paragraph_format.line_spacing = Pt(20)
    if indent_level > 0:
        p.paragraph_format.left_indent = Cm(0.74 * (indent_level + 1))
    return p


def add_numbered_step(doc, step_num, text, font_size=Pt(11)):
    """「步骤 N：」样式的操作步骤。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.74)
    pf.space_after = Pt(4)
    pf.line_spacing = Pt(20)
    add_run(p, f"步骤 {step_num}：", SONG, font_size, bold=True)
    add_rich_text(p, text, SONG, font_size)
    return p


def add_note(doc, text, note_type="提示"):
    """带底色的提示框，用单行无边框表格实现。"""
    palette = {
        "提示": ("E8F5E9", "1B5E20"),
        "注意": ("FFF3E0", "BF360C"),
        "警告": ("FFEBEE", "B71C1C"),
        "重要": ("E3F2FD", "0D47A1"),
    }
    bg, fg = palette.get(note_type, ("F5F5F5", "333333"))

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._tbl
    tblPr = _ensure_tblPr(tbl)

    _strip_table_borders(tblPr)
    _set_table_width(tblPr, "9000")

    cell = table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()

    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), bg)
    tcPr.append(shading)

    tcMar = OxmlElement("w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        mar = OxmlElement(f"w:{side}")
        mar.set(qn("w:w"), "120")
        mar.set(qn("w:type"), "dxa")
        tcMar.append(mar)
    tcPr.append(tcMar)

    cell.paragraphs[0].clear()
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.line_spacing = Pt(18)

    label_color = RGBColor(int(fg[0:2], 16), int(fg[2:4], 16), int(fg[4:6], 16))
    add_run(p, f"【{note_type}】", HEI, Pt(10), bold=True, color=label_color)
    add_rich_text(p, text, SONG, Pt(10), color=DARK)
    return table


def add_note_after_space(doc, text, note_type="提示"):
    """提示框，前后各留一点间距。"""
    add_paragraph(doc, "", font_size=Pt(4), space_after=Pt(0))
    table = add_note(doc, text, note_type)
    add_paragraph(doc, "", font_size=Pt(4), space_after=Pt(0))
    return table


def add_image_placeholder(doc, width_cm=15.0, height_cm=9.0,
                          label="【 请在此处插入截图 】"):
    """虚线框 + 灰底的截图占位区。"""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tbl = table._tbl
    tblPr = OxmlElement("w:tblPr")
    _set_table_width(tblPr, str(int(Emu(Cm(width_cm)) / 635)))

    borders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{name}")
        border.set(qn("w:val"), "dashed")
        border.set(qn("w:sz"), "10")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "999999")
        borders.append(border)
    tblPr.append(borders)

    old = tbl.find(qn("w:tblPr"))
    if old is not None:
        tbl.remove(old)
    tbl.insert(0, tblPr)

    row = table.rows[0]
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(Emu(Cm(height_cm)) / 635)))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)

    cell = table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)

    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), "F5F5F5")
    tcPr.append(shading)

    cell.paragraphs[0].clear()
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(6)
    add_run(p, label, SONG, Pt(13), bold=True,
            color=RGBColor(0x99, 0x99, 0x99))
    return table


def add_separator(doc, char="─", count=60):
    """一条浅灰分隔线。"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(4)
    add_run(p, char * count, SONG, Pt(6), color=RGBColor(0xCC, 0xCC, 0xCC))
    return p


def add_page_break(doc):
    doc.add_page_break()


def add_dotted_leader(doc, left, right, font_size=Pt(10.5)):
    """「条目 ........ 页码」这种带点线的行（手写目录用）。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = Pt(20)
    pf = p.paragraph_format
    pf.tab_stops.add_tab_stop(Cm(15.5), WD_ALIGN_PARAGRAPH.RIGHT)
    add_run(p, left, SONG, font_size)
    add_run(p, "\t", SONG, font_size)
    add_run(p, right, SONG, font_size)
    return p


# ──────────────────────────────────────────────
# 表格
# ──────────────────────────────────────────────

def add_grid_table(doc, headers: list[str], rows: list[list[str]],
                   widths_cm: list[float] | None = None,
                   header_fill="D2691E",
                   body_font_size=Pt(9.5),
                   header_font_size=Pt(10)):
    """带表头底色的网格表。表头文字为白色。"""
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        set_cell_text(cell, h, HEI, header_font_size, bold=True,
                      alignment=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, header_fill)
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for i, row in enumerate(rows, start=1):
        for j, value in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.CENTER if j == 0 else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(table.cell(i, j), value, SONG, body_font_size,
                          alignment=align)

    if widths_cm:
        set_col_widths(table, widths_cm)
    return table


# ──────────────────────────────────────────────
# 文档级
# ──────────────────────────────────────────────

def setup_page(doc, width_cm=21.0, height_cm=29.7,
               left_cm=2.0, right_cm=2.0, top_cm=1.8, bottom_cm=1.5,
               base_font=SONG, base_size=Pt(11), line_spacing=Pt(20)):
    """A4 页面 + 默认字体。"""
    section = doc.sections[0]
    section.page_width = Cm(width_cm)
    section.page_height = Cm(height_cm)
    section.left_margin = Cm(left_cm)
    section.right_margin = Cm(right_cm)
    section.top_margin = Cm(top_cm)
    section.bottom_margin = Cm(bottom_cm)

    style = doc.styles["Normal"]
    style.font.name = base_font
    style.font.size = base_size
    style.element.rPr.rFonts.set(qn("w:eastAsia"), base_font)
    style.paragraph_format.line_spacing = line_spacing
    return section


def setup_headings(doc, font_name=HEI):
    """把三级内置标题的字体统一成黑体。"""
    for level in (1, 2, 3):
        style = doc.styles[f"Heading {level}"]
        style.font.name = font_name
        style.element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def add_cover(doc, big_title, subtitle, info_lines,
              agency="计算机软件著作权登记", note=None,
              rule_char="━", rule_count=35):
    """封面：登记字样 + 大标题 + 副标题 + 信息行。"""
    for _ in range(5):
        add_paragraph(doc, "", font_size=Pt(12))

    add_paragraph(doc, agency, font_name=HEI, font_size=Pt(22), bold=True,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(16))
    add_paragraph(doc, big_title, font_name=HEI, font_size=Pt(26), bold=True,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(30))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, rule_char * rule_count, SONG, Pt(11), color=DARK)

    add_paragraph(doc, "", font_size=Pt(8))

    for line in info_lines:
        gray = line.startswith("（") or line.startswith("(")
        add_paragraph(doc, line, font_name=SONG, font_size=Pt(14),
                      alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(8),
                      color=GREY if gray else None)

    if note:
        add_paragraph(doc, "", font_size=Pt(16))
        add_paragraph(doc, note, font_name=SONG, font_size=Pt(11),
                      alignment=WD_ALIGN_PARAGRAPH.CENTER, color=GREY)


def add_toc_field(doc, title="目  录", levels="1-3"):
    """插入 Word 目录域。打开文档后右键「更新域」即可看到目录。"""
    add_page_break(doc)
    add_paragraph(doc, title, font_name=HEI, font_size=Pt(18), bold=True,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(16))
    add_paragraph(doc,
                  "（在 Word 中右键此处 → 更新域 → 更新整个目录，即可自动生成目录）",
                  font_name=SONG, font_size=Pt(9),
                  alignment=WD_ALIGN_PARAGRAPH.CENTER, color=GREY,
                  space_after=Pt(12))

    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)

    run2 = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f' TOC \\o "{levels}" \\h \\z \\u '
    run2._r.append(instr)

    run3 = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    run3._r.append(separate)

    run4 = paragraph.add_run("（请在 Word 中更新目录域）")
    run4.font.color.rgb = GREY
    run4.font.size = Pt(9)

    run5 = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run5._r.append(end)


# ──────────────────────────────────────────────
# 内部
# ──────────────────────────────────────────────

def _ensure_tblPr(tbl):
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    return tblPr


def _set_table_width(tblPr, dxa: str):
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), dxa)
    tblW.set(qn("w:type"), "dxa")
    old = tblPr.find(qn("w:tblW"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(tblW)


def _strip_table_borders(tblPr):
    borders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{name}")
        border.set(qn("w:val"), "none")
        border.set(qn("w:sz"), "0")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "auto")
        borders.append(border)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tblPr.append(borders)
