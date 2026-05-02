#!/usr/bin/env python3
"""
contextual — developed by @anshdhariwal
Run with: python contextual.py
"""

import sys, os, re, time, subprocess

def _importable(pkg):
    try: __import__(pkg); return True
    except ImportError: return False

_needed = {"pdfplumber": "pdfplumber", "pypdf": "pypdf",
           "docx": "python-docx", "pptx": "python-pptx"}
_missing = [pip_name for mod, pip_name in _needed.items() if not _importable(mod)]
if _missing:
    print(f"\n  Installing: {', '.join(_missing)} …\n")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet"] + _missing)

import pdfplumber
from pypdf import PdfReader
import docx as _docx
from pptx import Presentation as _Presentation

def numeric_key(name: str):
    m = re.match(r"^(\d+)", name)
    return (int(m.group(1)), name) if m else (float("inf"), name)

def discover_files() -> list:
    exts = (".pdf", ".docx", ".pptx")
    return sorted(
        [f for f in os.listdir(".") if f.lower().endswith(exts)],
        key=numeric_key,
    )

def extract_pdf(path: str, page_nums: bool) -> tuple:
    for fn, name in (
        (lambda: "\n\n".join(
            (f"[Page {i}]\n" if page_nums else "") + (p.extract_text() or '')
            for i, p in enumerate(pdfplumber.open(path).pages, 1)),
         "pdfplumber"),
        (lambda: "\n\n".join(
            (f"[Page {i}]\n" if page_nums else "") + (p.extract_text() or '')
            for i, p in enumerate(PdfReader(path).pages, 1)),
         "pypdf"),
    ):
        try:
            t = fn()
            if t.strip(): return t, name
        except Exception:
            pass
    return "[EXTRACTION ERROR]", "error"

def extract_docx(path: str, page_nums: bool) -> tuple:
    try:
        doc = _docx.Document(path)
        text = "\n".join(p.text for p in doc.paragraphs)
        return text, "python-docx"
    except Exception as e:
        return f"[EXTRACTION ERROR — {e}]", "error"

def extract_pptx(path: str, page_nums: bool) -> tuple:
    try:
        prs = _Presentation(path)
        slides = []
        for i, slide in enumerate(prs.slides, 1):
            texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        texts.append(para.text)
            slides.append((f"[Slide {i}]\n" if page_nums else "") + "\n".join(texts))
        return "\n\n".join(slides), "python-pptx"
    except Exception as e:
        return f"[EXTRACTION ERROR — {e}]", "error"

def extract(path: str, page_nums: bool) -> tuple:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx": return extract_docx(path, page_nums)
    if ext == ".pptx": return extract_pptx(path, page_nums)
    return extract_pdf(path, page_nums)
   
def main():
    print("\n              [ contextual — @anshdhariwal ]\n")
    print("  tool uses all .pdf, .docx and .pptx files in the current directory\n")

    files = discover_files()

    if not files:
        print("  No .pdf, .docx, or .pptx files found here.")
        print(f"  CWD: {os.getcwd()}")
        sys.exit(1)

    print(f"  Found {len(files)} file(s) in: {os.getcwd()}\n")
    print("  >   FIles included:")
    print("  " + "─" * 40)
    for i, f in enumerate(files, 1):
        print(f"  {i:<4}{f}")
    print()

    input("  Press Enter to continue … ")
    print()

    while True:
        name = input(" > Enter Output filename (without .txt): ").strip()
        if name:
            name = name.removesuffix(".txt")
            break
        print("  Filename cannot be empty.\n")

    while True:
        pn = input(" > Number pages/slides in output? (y/n): ").strip().lower()
        if pn in ("y", "n"): break
        print("  Type y or n.\n")
    page_nums = pn == "y"
    print()

    print("  making your file buddy………")

    start = time.time()

    DIV = "─" * 72

    results = []
    for f in files:
        text, engine = extract(f, page_nums)
        results.append((f, text, engine))

    lines = [
        DIV,
        f"  CONCAT OF {len(files)} FILE{'S' if len(files) != 1 else ''}",
        f"  contextual  ·  @anshdhariwal",
        f"  Generated : {time.strftime('%Y-%m-%d  %H:%M:%S')}",
        DIV, "",
        "  TABLE OF CONTENTS", DIV,
    ]
    for i, (f, _, _) in enumerate(results, 1):
        lines.append(f"    {i:>3}.  {f}")
    lines += ["", DIV, ""]

    for i, (f, text, engine) in enumerate(results, 1):
        lines += [
            f"  [{i}/{len(results)}]  {f}",
            f"  engine : {engine}",
            DIV, "",
            text.strip() if text.strip() else "[No text extracted]",
            "", DIV, "",
        ]

    out_path = f"{name}.txt"
    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))

    elapsed = time.time() - start
    if elapsed < 2:
        time.sleep(2 - elapsed)

    print()
    print(f"  ✓  Saved to {out_path}")
    print(f"     {os.path.abspath(out_path)}")
    print()
    print("  work done, good bye!")
    time.sleep(3)

if __name__ == "__main__":
    main()
