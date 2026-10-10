# -*- coding: utf-8 -*-
"""对渲染结果做程序化体检：空页、文字/图片越界、页面填充率。"""
import os, sys, io, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import fitz

HERE = os.path.dirname(os.path.abspath(__file__))
pdfs = glob.glob(os.path.join(HERE, "_render", "*.pdf"))
doc = fitz.open(pdfs[0])
print("pages:", doc.page_count)

issues = []
low_fill = []
for i, page in enumerate(doc, 1):
    rect = page.rect
    txt = "".join(page.get_text().split())
    if len(txt) < 12 and i != 1:
        issues.append("P%02d 近似空白页（字符数=%d）" % (i, len(txt)))

    # 越界检测
    for b in page.get_text("blocks"):
        x0, y0, x1, y1 = b[:4]
        if x0 < 2 or y0 < 2 or x1 > rect.width - 2 or y1 > rect.height - 2:
            issues.append("P%02d 文本越界 (%.1f,%.1f)-(%.1f,%.1f) 「%s」"
                          % (i, x0, y0, x1, y1, "".join(b[4].split())[:20]))
            break
    for img in page.get_image_info():
        bb = img["bbox"]
        if bb[0] < 2 or bb[2] > rect.width - 2:
            issues.append("P%02d 图片越界 %s" % (i, [round(v, 1) for v in bb]))
            break

    # 填充率：正文块（排除页眉页脚）的纵向覆盖比例
    body = [b for b in page.get_text("blocks")
            if b[1] > 60 and b[3] < rect.height - 60]
    if body:
        top = min(b[1] for b in body)
        bot = max(b[3] for b in body)
        # 图片也计入
        for img in page.get_image_info():
            top = min(top, img["bbox"][1])
            bot = max(bot, img["bbox"][3])
        usable = rect.height - 60 * 2
        fill = (bot - top) / usable
        if fill < 0.42 and i != 1:
            low_fill.append((i, round(fill * 100)))

if issues:
    print("ISSUES (%d):" % len(issues))
    for s in issues:
        print("  -", s)
else:
    print("NO OVERFLOW / BLANK ISSUES")

if low_fill:
    print("LOW FILL (<42%%) — 页面偏空 (%d):" % len(low_fill))
    for p, f in low_fill:
        print("  P%02d  fill=%d%%" % (p, f))
else:
    print("no low-fill pages")
