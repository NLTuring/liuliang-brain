# -*- coding: utf-8 -*-
"""将作品书 docx 渲染为 PDF，再逐页导出 PNG，用于视觉验收。"""
import os, sys, io, glob, subprocess, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(HERE)
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
OUTDIR = os.path.join(HERE, "_render")
PAGEDIR = os.path.join(HERE, "_pages")

docx_path = sys.argv[1] if len(sys.argv) > 1 else None
if not docx_path:
    cands = [f for f in os.listdir(PROJ) if f.endswith(".docx") and "作品书" in f]
    docx_path = os.path.join(PROJ, cands[0])
print("input:", docx_path)

if os.path.isdir(OUTDIR):
    shutil.rmtree(OUTDIR)
os.makedirs(OUTDIR, exist_ok=True)
if os.path.isdir(PAGEDIR):
    shutil.rmtree(PAGEDIR)
os.makedirs(PAGEDIR, exist_ok=True)

cmd = [SOFFICE, "--headless", "--norestore", "--convert-to", "pdf",
       "--outdir", OUTDIR, docx_path]
print("running:", " ".join(cmd))
r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
print("rc:", r.returncode)
print(r.stdout[-2000:])
print(r.stderr[-2000:])

pdfs = glob.glob(os.path.join(OUTDIR, "*.pdf"))
if not pdfs:
    print("!! PDF not produced")
    sys.exit(1)
pdf = pdfs[0]
print("pdf:", pdf, os.path.getsize(pdf), "bytes")

import fitz
doc = fitz.open(pdf)
print("pages:", doc.page_count)
for i, page in enumerate(doc, 1):
    pix = page.get_pixmap(dpi=110)
    out = os.path.join(PAGEDIR, "page-%02d.png" % i)
    pix.save(out)
print("rendered pages ->", PAGEDIR)
