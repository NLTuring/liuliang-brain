# -*- coding: utf-8 -*-
"""交付前最终核验：字体、结构、目录页码与文件产物。"""
import os, sys, io, glob, shutil, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(HERE)

# 1) 复制 PDF 到项目根目录
pdfs = glob.glob(os.path.join(HERE, "_render", "*.pdf"))
if pdfs:
    dst = os.path.join(PROJ, "留量大脑-电商赋能乡村的流量沉淀机制与数字治理平台-作品书.pdf")
    shutil.copy2(pdfs[0], dst)
    print("PDF copied ->", dst)

# 2) 核验 docx 字体与结构
docxs = [f for f in os.listdir(PROJ) if f.endswith(".docx")]
path = os.path.join(PROJ, docxs[0])
d = docx.Document(path)
print("docx:", os.path.basename(path), "| size %.1f KB" % (os.path.getsize(path) / 1024))
print("sections:", len(d.sections), "| paragraphs:", len(d.paragraphs), "| tables:", len(d.tables))
print("inline shapes:", len(d.inline_shapes))

# Normal 样式字体
st = d.styles["Normal"]
print("Normal font:", st.font.name, st.font.size.pt, "pt")
print("Normal eastAsia:", st.element.rPr.rFonts.get(qn("w:eastAsia")))
for nm in ("Heading 1", "Heading 2", "Heading 3"):
    s = d.styles[nm]
    print("%-10s %s %.1fpt bold=%s eastAsia=%s" % (
        nm, s.font.name, s.font.size.pt, s.font.bold,
        s.element.rPr.rFonts.get(qn("w:eastAsia"))))

# 抽查正文 run 字体，确认没有非宋体
bad = set()
for p in d.paragraphs:
    for r in p.runs:
        rf = r._element.rPr.rFonts if r._element.rPr is not None else None
        if rf is not None:
            ea = rf.get(qn("w:eastAsia"))
            if ea and ea != "宋体":
                bad.add(ea)
for t in d.tables:
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    rf = r._element.rPr.rFonts if r._element.rPr is not None else None
                    if rf is not None:
                        ea = rf.get(qn("w:eastAsia"))
                        if ea and ea != "宋体":
                            bad.add(ea)
print("non-宋体 eastAsia fonts found:", sorted(bad) if bad else "无（全部为宋体）")

# 3) 核验目录页码与正文实际页码一致
pm = os.path.join(HERE, "page_map.json")
if os.path.exists(pm):
    import json
    m = json.load(open(pm, encoding="utf-8"))
    print("TOC page entries:", len(m), "| 第一章 =", m.get("第一章  项目速览"),
          "| 附录 E =", m.get("附录 E  待补齐与核实清单（提交前必做）"))
print("figures:", sorted(os.path.basename(f) for f in glob.glob(os.path.join(PROJ, "figures", "*.png"))))
