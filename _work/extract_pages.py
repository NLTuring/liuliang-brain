# -*- coding: utf-8 -*-
"""从渲染出的 PDF 中提取各标题所在页码，写入 page_map.json。

匹配策略：按“行文本 + 字号”双重匹配，避免短标题（如“对商户”）在正文中被
误匹配到不相干的页面。标题字号由 build 时的 H_SIZES 决定。
"""
import os, sys, io, json, re, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fitz
import build_docx as B
from docx_common import H_SIZES

TOL = 0.9


def norm(s):
    return re.sub(r"\s+", "", s)


def main():
    pdfs = glob.glob(os.path.join(HERE, "_render", "*.pdf"))
    if not pdfs:
        print("no pdf found; run render.py first")
        return 1
    doc = fitz.open(pdfs[0])
    pages = [p.get_text() for p in doc]

    BODY_MARK = norm("本章是全书内容的浓缩版")
    body_start = 0
    for i, t in enumerate(pages):
        if BODY_MARK in norm(t):
            body_start = i
            break
    print("body starts at PDF page 1-based:", body_start + 1)

    headings = B.collect_headings()
    hmap = {}
    for lvl, text in headings:
        hmap.setdefault(norm(text), lvl)
    print("headings:", len(headings), "unique:", len(hmap))

    assigned = {}
    size_seen = {}
    for i in range(body_start, doc.page_count):
        d = doc[i].get_text("dict")
        for blk in d.get("blocks", []):
            for line in blk.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                lt = norm("".join(s["text"] for s in spans))
                if not lt:
                    continue
                for s in spans:
                    key = round(s["size"], 1)
                    size_seen[key] = size_seen.get(key, 0) + 1
                if lt in hmap and lt not in assigned:
                    lvl = hmap[lt]
                    mx = max(s["size"] for s in spans)
                    if abs(mx - H_SIZES[lvl]) < TOL:
                        assigned[lt] = i + 1

    print("font sizes seen (size: count):",
          dict(sorted(size_seen.items(), key=lambda kv: -kv[1])[:10]))

    missing = []
    pmap = {}
    for lvl, text in headings:
        nt = norm(text)
        if nt in assigned:
            pmap[text] = assigned[nt]
        else:
            # 回退：正文页内子串匹配（可能不准，会告警）
            found = None
            for i in range(body_start, doc.page_count):
                if nt in norm(pages[i]):
                    found = i + 1
                    break
            if found is None:
                missing.append(text)
            else:
                pmap[text] = found
                missing.append("(fallback) " + text)

    if missing:
        print("WARNING %d headings not matched by size:" % len(missing))
        for m in missing:
            print("   -", m)

    out = os.path.join(HERE, "page_map.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(pmap, f, ensure_ascii=False, indent=1)
    print("wrote", out, "entries:", len(pmap))
    for lvl, text in headings:
        if lvl == 1:
            print("   %-34s -> %s" % (text, pmap.get(text, "?")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
