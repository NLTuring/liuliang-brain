# -*- coding: utf-8 -*-
"""一键流水线：构建 → 渲染 → 提取页码 → 再构建（回填页码）→ 渲染 → 体检。

页码回填需要两次构建，因为目录页码要等第一次渲染后才能知道。
"""
import os, subprocess, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

def run(script):
    print("\n=== %s ===" % script)
    r = subprocess.run([PY, os.path.join(HERE, script)], cwd=HERE)
    if r.returncode != 0:
        print("!! %s failed rc=%s" % (script, r.returncode))
        sys.exit(r.returncode)

if __name__ == "__main__":
    pm = os.path.join(HERE, "page_map.json")
    run("build_docx.py")
    run("render.py")
    run("extract_pages.py")
    run("build_docx.py")
    run("render.py")
    run("extract_pages.py")
    run("check_pages.py")
    print("\nPIPELINE DONE")
