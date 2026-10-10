# -*- coding: utf-8 -*-
"""《留量大脑》作品书排版工具。

统一字体：宋体（SimSun），中英文均使用宋体；标题加粗。
版式：A4，版心约 15.8 × 22.0 cm，正文五号（10.5pt），行距 1.45。
"""
import os
import re
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "宋体"

ACCENT = RGBColor(0x1B, 0x4D, 0x3E)      # 墨绿：一级标题
ACCENT2 = RGBColor(0x2E, 0x7D, 0x64)     # 中绿：二/三级标题
DARK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x59, 0x59, 0x59)
WARN = RGBColor(0xA8, 0x3A, 0x32)

HEADER_FILL = "1B4D3E"
ZEBRA_FILL = "EDF5F1"
CALLOUT_FILL = "EAF4EF"
QUOTE_FILL = "F2F7F4"

PAGE_W_CM = 21.0
PAGE_H_CM = 29.7
MARGIN_LR_CM = 2.6
MARGIN_TB_CM = 2.54
TABLE_TOTAL_CM = 15.8

BODY_PT = 10.5
H_SIZES = {1: 16.0, 2: 13.5, 3: 12.0}
H_COLORS = {1: ACCENT, 2: ACCENT2, 3: ACCENT2}


# ------------------------------------------------------------------ 底层工具
def set_run(run, size=BODY_PT, bold=False, color=None, italic=False, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color if color is not None else DARK
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), font)
    return run


def shade_element(element, fill):
    pr = element.get_or_add_tcPr() if element.tag.endswith("}tc") else element.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pr.append(shd)


def cell_shade(cell, fill):
    shade_element(cell._tc, fill)


def para_shade(paragraph, fill):
    shade_element(paragraph._p, fill)


def para_left_border(paragraph, color="2E7D64", size=18):
    pPr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color)
    borders.append(left)
    pPr.append(borders)


def cell_margins(table, top=60, bottom=60, left=95, right=95):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement("w:" + tag)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        mar.append(node)
    tblPr.append(mar)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:cantSplit")
    node.set(qn("w:val"), "true")
    trPr.append(node)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    trPr.append(node)


def keep_with_next(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:keepNext"))


def set_table_widths(t, widths_cm):
    """显式写入 tblGrid / tblW / tblLayout，避免渲染器把表格拉成等宽而溢出。"""
    tbl = t._tbl
    tblPr = tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        el = tblPr.find(qn(tag))
        if el is not None:
            tblPr.remove(el)

    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(int(round(sum(widths_cm) * 567))))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(round(w * 567))))
        grid.append(gc)
    tbl.insert(list(tbl).index(tblPr) + 1, grid)

    for row in t.rows:
        for i, w in enumerate(widths_cm):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def fit_widths(widths):
    total = float(sum(widths))
    if total <= 0:
        return list(widths)
    k = TABLE_TOTAL_CM / total
    return [w * k for w in widths]


# ------------------------------------------------------------------ 样式
def init_styles(doc):
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(BODY_PT)
    st.font.color.rgb = DARK
    for a in ("w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"):
        st.element.rPr.rFonts.set(qn(a), FONT)
    st.paragraph_format.line_spacing = 1.45
    st.paragraph_format.space_after = Pt(4)

    for name, size, color, before, after in (
        ("Heading 1", H_SIZES[1], H_COLORS[1], 16, 9),
        ("Heading 2", H_SIZES[2], H_COLORS[2], 12, 6),
        ("Heading 3", H_SIZES[3], H_COLORS[3], 9, 4),
    ):
        s = doc.styles[name]
        s.font.name = FONT
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = color
        for a in ("w:eastAsia", "w:ascii", "w:hAnsi", "w:cs"):
            s.element.rPr.rFonts.set(qn(a), FONT)
        s.paragraph_format.space_before = Pt(before)
        s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.line_spacing = 1.3
        s.paragraph_format.keep_with_next = True


def setup_page(sec):
    sec.page_width = Cm(PAGE_W_CM)
    sec.page_height = Cm(PAGE_H_CM)
    sec.top_margin = Cm(MARGIN_TB_CM)
    sec.bottom_margin = Cm(MARGIN_TB_CM)
    sec.left_margin = Cm(MARGIN_LR_CM)
    sec.right_margin = Cm(MARGIN_LR_CM)


def _field(paragraph, instr, placeholder="1", size=9, color=None, font=FONT):
    """向段落追加一个简单域（如 PAGE / TOC）。"""
    run = paragraph.add_run()
    r = run._r
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = placeholder
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    for el in (fld_begin, instr_el, fld_sep, t, fld_end):
        r.append(el)
    set_run(run, size, color=color, font=font)
    return run


def add_body_footer(sec):
    """正文区页脚：第 X 页。"""
    sec.footer.is_linked_to_previous = False
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("第 "), 9, color=GREY)
    _field(p, "PAGE", "1", 9, GREY)
    set_run(p.add_run(" 页"), 9, color=GREY)


def add_body_header(sec, text):
    sec.header.is_linked_to_previous = False
    p = sec.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(text), 8.5, color=GREY)
    pPr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "BFBFBF")
    borders.append(bottom)
    pPr.append(borders)


def add_toc_field(doc, entries, levels="1-3", right_cm=15.8):
    """写入带缓存条目的 TOC 域。

    entries: [(level, text, page_str), ...]；page_str 用于域名未更新时的显示，
    Word / WPS 打开即可看到目录，之后仍可右键“更新域”刷新。
    """
    if not entries:
        return None

    paras = []
    for idx, (lvl, text, page) in enumerate(entries):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.0 + 0.66 * (lvl - 1))
        pf.right_indent = Cm(0.0)
        pf.space_before = Pt(0.8 if lvl == 1 else 0)
        pf.space_after = Pt(0.2)
        pf.line_spacing = 1.0
        pPr = p._p.get_or_add_pPr()
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), str(int(round(right_cm * 567))))
        tabs.append(tab)
        pPr.append(tabs)

        # 首段：先写入域名开始标记，再写可见条目文本（保证 field 包裹整段目录）
        if idx == 0:
            start = p.add_run()
            set_run(start, 9)
            b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
            i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve")
            i.text = 'TOC \\o "%s" \\h \\z \\u' % levels
            s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate")
            start._r.append(b); start._r.append(i); start._r.append(s)

        size = {1: 10.5, 2: 9.5, 3: 9.0}[lvl]
        color = ACCENT if lvl == 1 else (DARK if lvl == 2 else GREY)
        set_run(p.add_run(text), size, bold=(lvl == 1), color=color)
        set_run(p.add_run("\t" + str(page)), size, bold=(lvl == 1), color=color)

        # 末段：收尾标记
        if idx == len(entries) - 1:
            end = p.add_run()
            set_run(end, 9)
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
            end._r.append(e)
        paras.append(p)
    return paras


# ------------------------------------------------------------------ 构建块
def H(doc, text, level, page_break_before=False):
    p = doc.add_paragraph(style="Heading %d" % level)
    if page_break_before:
        p.paragraph_format.page_break_before = True
    set_run(p.add_run(text), H_SIZES[level], True, H_COLORS[level])
    return p


def add_rich(paragraph, text, size=BODY_PT, color=None, bold=False):
    """把 **粗体** 解析为真正的加粗 run。"""
    pos = 0
    for m in re.finditer(r"\*\*(.+?)\*\*", text, re.S):
        if m.start() > pos:
            set_run(paragraph.add_run(text[pos:m.start()]), size, bold=bold, color=color)
        set_run(paragraph.add_run(m.group(1)), size, bold=True,
                color=color if color is not None else ACCENT)
        pos = m.end()
    if pos < len(text):
        set_run(paragraph.add_run(text[pos:]), size, bold=bold, color=color)
    return paragraph


def keep_together(paragraph):
    """段中不分页：避免一两个字的尾巴被挤到下一页，形成近乎空白的页面。"""
    pPr = paragraph._p.get_or_add_pPr()
    node = OxmlElement("w:keepLines")
    node.set(qn("w:val"), "true")
    pPr.append(node)


def para(doc, text, indent=True, size=BODY_PT, align=None, after=4, color=None):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    p.paragraph_format.space_after = Pt(after)
    if align is not None:
        p.alignment = align
    keep_together(p)
    if "**" in text:
        add_rich(p, text, size, color=color)
    else:
        set_run(p.add_run(text), size, color=color)
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.7 + 0.62 * level)
    p.paragraph_format.first_line_indent = Cm(-0.52)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.4
    mark = "● " if level == 0 else "○ "
    keep_together(p)
    set_run(p.add_run(mark), 8, color=ACCENT2)
    if "**" in text:
        add_rich(p, text)
    else:
        set_run(p.add_run(text), BODY_PT)
    return p


def numbered(doc, items):
    for i, text in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.85)
        p.paragraph_format.first_line_indent = Cm(-0.85)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.42
        keep_together(p)
        set_run(p.add_run("%d. " % i), BODY_PT, bold=True, color=ACCENT2)
        if "**" in text:
            add_rich(p, text)
        else:
            set_run(p.add_run(text), BODY_PT)
    return None


def quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.7)
    p.paragraph_format.right_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.4
    para_shade(p, QUOTE_FILL)
    para_left_border(p)
    keep_together(p)
    if "**" in text:
        add_rich(p, text, BODY_PT)
    else:
        set_run(p.add_run(text), BODY_PT)
    return p


def table(doc, headers, rows, widths=None, size=9.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_margins(t)

    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        set_run(p.add_run(h), size, True, RGBColor(0xFF, 0xFF, 0xFF))
        cell_shade(c, HEADER_FILL)
    repeat_header(hdr)
    no_split(hdr)

    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        no_split(t.rows[-1])
        for i, val in enumerate(row):
            c = cells[i]
            c.text = ""
            for li, line in enumerate(str(val).split("\n")):
                p = c.paragraphs[0] if li == 0 else c.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.2
                if "**" in line:
                    add_rich(p, line, size)
                else:
                    set_run(p.add_run(line), size)
            if ri % 2 == 1:
                cell_shade(c, ZEBRA_FILL)

    if widths:
        set_table_widths(t, fit_widths(widths))
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 1.0
    set_run(sp.add_run(""), 4)
    return t


def callout(doc, title_text, lines):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.autofit = False
    cell = t.rows[0].cells[0]
    cell.text = ""
    cell_margins(t, top=130, bottom=130, left=170, right=170)
    cell_shade(cell, CALLOUT_FILL)
    no_split(t.rows[-1])
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    set_run(p.add_run(title_text), 11.5, True, ACCENT)
    for line in lines:
        q = cell.add_paragraph()
        q.paragraph_format.space_after = Pt(3)
        q.paragraph_format.line_spacing = 1.4
        q.paragraph_format.first_line_indent = Pt(21)
        if "**" in line:
            add_rich(q, line, BODY_PT)
        else:
            set_run(q.add_run(line), BODY_PT)
    set_table_widths(t, [TABLE_TOTAL_CM])
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 1.0
    set_run(sp.add_run(""), 4)
    return t


def caption(doc, text, keep_next=False, as_title=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(3 if as_title else 8)
    p.paragraph_format.line_spacing = 1.25
    if as_title:
        p.paragraph_format.space_before = Pt(2)
    if keep_next:
        keep_with_next(p)
    color = RGBColor(0x33, 0x33, 0x33) if as_title else GREY
    size = 9.0 if as_title else 8.5
    if "**" in text:
        add_rich(p, text, size, color=color)
    else:
        set_run(p.add_run(text), size, color=color)
    return p


def figure(doc, path, cap, width_cm=15.0, figdir=""):
    full = path if os.path.isabs(path) else os.path.join(figdir, path)
    if not os.path.exists(full):
        print("  ! figure missing:", full)
        return None
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    p.add_run().add_picture(full, width=Cm(width_cm))
    keep_with_next(p)
    q = doc.add_paragraph()
    q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    q.paragraph_format.space_after = Pt(9)
    q.paragraph_format.line_spacing = 1.25
    set_run(q.add_run(cap), 8.5, color=GREY)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)
    return p


def new_document():
    doc = Document()
    init_styles(doc)
    setup_page(doc.sections[0])
    return doc
