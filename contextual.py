#!/usr/bin/env python3
"""
contextual — developed by @anshdhariwal
Run with: python contextual.py
"""

import sys, os, re, time, subprocess, io, hashlib

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
from docx.oxml.ns import qn as _qn
from pptx import Presentation as _Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image

OCR_MIN_BYTES = 10 * 1024
OCR_MIN_SIDE = 64
OCR_MAX_SIDE = 2000

_ocr = False   # False = not loaded yet, None = unavailable, else the engine
_ocr_cache = {}

def get_ocr():
    global _ocr
    if _ocr is not False:
        return _ocr
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        print("\n  Installing OCR engine (rapidocr-onnxruntime, one-time) …")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install",
                                   "--quiet", "rapidocr-onnxruntime"])
            from rapidocr_onnxruntime import RapidOCR
        except Exception:
            print("  couldn't set up OCR — continuing without it.\n")
            _ocr = None
            return None
    print("  loading OCR model …")
    _ocr = RapidOCR()
    return _ocr

def ocr_blob(blob: bytes) -> str:
    key = hashlib.md5(blob).hexdigest()
    if key in _ocr_cache:
        return _ocr_cache[key]
    text = ""
    img = None
    try:
        img = Image.open(io.BytesIO(blob)).convert("RGB")
        w, h = img.size
        if len(blob) < OCR_MIN_BYTES or w < OCR_MIN_SIDE or h < OCR_MIN_SIDE:
            img = None
    except Exception:
        img = None
    if img is not None:
        engine = get_ocr()
        if engine is not None:
            try:
                if max(w, h) > OCR_MAX_SIDE:
                    scale = OCR_MAX_SIDE / max(w, h)
                    img = img.resize((int(w * scale), int(h * scale)))
                import numpy as np
                result, _ = engine(np.array(img)[:, :, ::-1])
                if result:
                    result.sort(key=lambda r: r[0][0][1])
                    text = "\n".join(t for _, t, s in result)
            except Exception:
                text = ""
    _ocr_cache[key] = text
    return text

def numeric_key(name: str):
    m = re.match(r"^(\d+)", name)
    return (int(m.group(1)), name) if m else (float("inf"), name)

def discover_files() -> list:
    exts = (".pdf", ".docx", ".pptx")
    return sorted(
        [f for f in os.listdir(".") if f.lower().endswith(exts)],
        key=numeric_key,
    )

def _extract_pdf_ocr(path: str, page_nums: bool) -> tuple:
    import fitz
    pages = []
    with fitz.open(path) as pdf:
        for i, page in enumerate(pdf, 1):
            items = []
            for b in page.get_text("dict")["blocks"]:
                y = b["bbox"][1]
                if b["type"] == 0:
                    txt = "\n".join("".join(s["text"] for s in ln["spans"])
                                    for ln in b["lines"])
                    if txt.strip():
                        items.append((y, 0, txt))
                else:
                    found = ocr_blob(b.get("image") or b"")
                    if found:
                        items.append((y, 1, f"[Image]\n{found}"))
            items.sort(key=lambda it: (it[0], it[1]))
            body = "\n\n".join(t for _, _, t in items)
            pages.append((f"[Page {i}]\n" if page_nums else "") + body)
    text = "\n\n".join(pages)
    return (text, "pymupdf+ocr")

def extract_pdf(path: str, page_nums: bool, ocr: bool = False) -> tuple:
    if ocr:
        try:
            return _extract_pdf_ocr(path, page_nums)
        except Exception:
            pass
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

def _para_images(p) -> list:
    blobs = []
    for blip in p._p.iter(_qn("a:blip")):
        rid = blip.get(_qn("r:embed"))
        try:
            blobs.append(p.part.rels[rid].target_part.blob)
        except Exception:
            pass
    return blobs

def extract_docx(path: str, page_nums: bool, ocr: bool = False) -> tuple:
    try:
        doc = _docx.Document(path)
        items = []
        for p in doc.paragraphs:
            t = p.text
            if ocr:
                for blob in _para_images(p):
                    found = ocr_blob(blob)
                    if found:
                        t += f"\n[Image]\n{found}"
            items.append((p, t))
        if page_nums:
            pages, cur = [], []
            for p, t in items:
                if any(run.text == "" and "w:lastRenderedPageBreak" in
                       (run._r.xml if hasattr(run, "_r") else "") for run in p.runs):
                    pages.append(cur); cur = []
                cur.append(t)
            pages.append(cur)
            text = "\n\n".join(f"[Page {i}]\n" + "\n".join(pg)
                               for i, pg in enumerate(pages, 1) if any(pg))
            return text, "python-docx"
        return "\n".join(t for _, t in items), "python-docx"
    except Exception as e:
        return f"[EXTRACTION ERROR — {e}]", "error"

def _slide_texts(shapes, ocr: bool) -> list:
    out = []
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            out += _slide_texts(sh.shapes, ocr)
            continue
        if ocr and sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.PLACEHOLDER):
            try:
                found = ocr_blob(sh.image.blob)
                if found:
                    out.append(f"[Image]\n{found}")
            except Exception:
                pass
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                out.append(para.text)
    return out

def extract_pptx(path: str, page_nums: bool, ocr: bool = False) -> tuple:
    try:
        prs = _Presentation(path)
        slides = []
        for i, slide in enumerate(prs.slides, 1):
            texts = _slide_texts(slide.shapes, ocr)
            slides.append((f"[Slide {i}]\n" if page_nums else "") + "\n".join(texts))
        return "\n\n".join(slides), "python-pptx"
    except Exception as e:
        return f"[EXTRACTION ERROR — {e}]", "error"

def extract(path: str, page_nums: bool, ocr: bool = False) -> tuple:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx": return extract_docx(path, page_nums, ocr)
    if ext == ".pptx": return extract_pptx(path, page_nums, ocr)
    return extract_pdf(path, page_nums, ocr)
   
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

    while True:
        oc = input(" > OCR images inside documents? (y/n): ").strip().lower()
        if oc in ("y", "n"): break
        print("  Type y or n.\n")
    use_ocr = oc == "y"
    print()

    print("  making your file buddy………")

    start = time.time()

    DIV = "─" * 72

    results = []
    for f in files:
        text, engine = extract(f, page_nums, use_ocr)
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
