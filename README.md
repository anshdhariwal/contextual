# contextual 📄

[![npm](https://img.shields.io/npm/v/contextual-py)](https://www.npmjs.com/package/contextual-py)
[![license](https://img.shields.io/npm/l/contextual-py)](LICENSE)

> your documents walked into a tool. they left as one.

a python CLI script that takes every PDF, DOCX, and PPTX in your current directory and converts them into a single, clean .txt file. no config, no flags, no extra hustle to play with. images inside your docs get read too - optional offline OCR, nothing leaves your machine.

---

## why?

because copy-pasting 37 files into one doc file at 2am for a context file is **not it**.
and, i felt the need for this be developed :)
contextual does the boring part so you can focus on your work.

---

## how it works

```
you run it → it finds your files → you press enter → you name the output → maybe say y to ocr → done.
```

that's it. that's the tool. its intentionally small and built for this specific purpose.

---

## quick start

no clone needed, if you have node:

```bash
npx contextual-py
```

or install it once and just run `contextual` anywhere:

```bash
npm i -g contextual-py
cd whatever-folder
contextual
```

old school also works:

```bash
cd whatever-folder
python contextual.py
```

### flags

every question can be pre-answered. answer everything and it runs with zero prompts:

```bash
contextual -f                          # merge all, output.txt, pages on, ocr off
contextual -o notes --no-pages --ocr   # custom name, one block per file, read images too
contextual --ocr                       # just preset ocr, rest still asks nicely
```

---

## requirements

- python 3.10+ (that's the actual engine)
- node 14+ (only for the npx/npm route)
- a folder with documents in it
- the will to merge them

everything else gets installed automatically on first run.

---

## features

| thing | status |
|---|---|
| PDF extraction | ✓ (pdfplumber + pypdf fallback) |
| DOCX extraction | ✓ (python-docx) |
| PPTX extraction | ✓ (python-pptx) |
| embedded image OCR | ✓ (offline, optional - you choose y/n) |
| auto-installs deps | ✓ (you literally just run it) |
| simple file sorting | ✓ (files starting with numbers sort correctly) |
| page/slide labels | ✓ (optional - you choose y/n) |
| table of contents in output | ✓ |

---

## the output

your `.txt` file comes with:
- a header with file count + timestamp
- a table of contents
- each file's content clearly separated with dividers
- page/slide numbers preserved

basically it looks like someone organized your clutter. thank me later.

## use cases

- **sending context to AI** - merge all your project docs into one file and feed it to ChatGPT, Claude, Gemini, or whatever LLM you're prompting. one file, full context.
- **studying for exams** - dump all your lecture slides and notes into a single searchable text file.
- **project documentation** - combine scattered reports, specs, and presentations into one reference doc before a meeting or deadline.
- **code review prep** - merge requirement docs, design specs, and test plans into a single file so reviewers have everything in one place.
- **archiving** - turn a folder of mixed-format documents into a single plain text backup that'll open on literally anything, forever.
- **you can** also pair this with markdown formatting for even better results.
---

## faq- you may ask

**Q: does it support xlsx?**
A: no. this is contextual, not excel anonymous.

**Q: can i reorder the files?**
A: they sort by leading number automatically. name them `1-intro.pdf`, `2-body.docx`, etc. or don't. i'm not your mom.

**Q: it says "making your file buddy" and waits 2 seconds even though it's already done.**
A: vibes. it's called pacing. just for ui.

**Q: why doesn't page numbering work for .docx files?**
A: because `.docx` files don't actually have "pages" as they're flow documents. page breaks only exist when Microsoft Word renders them on screen, and that info isn't stored in the file itself. so there's nothing to number. PDFs have real pages, PPTXs have real slides, DOCX has paragraphs. so, we just dump them all.

**Q: how does the OCR thing work?**
A: say `y` when asked, and any images embedded in your docs get read locally with rapidocr (a small onnx model, no cloud, no api keys). the text lands exactly where the image was. logos and icons are skipped automatically, and repeated images (like a logo on every slide) only get processed once. first `y` installs the engine one time, after that it's fully offline.

**Q: does OCR slow it down?**
A: only for files that actually contain readable images. clean text docs process at the same speed as before.

---

## changelog

- **v1.2** - cli flags. `-f` for fully unattended runs, or pre-answer any subset (`-o`, `--pages/--no-pages`, `--ocr/--no-ocr`) and let the rest still ask.
- **v1.1** - optional offline OCR for images inside PDFs, DOCX and PPTX. scanned PDFs work now too. also published on npm as [`contextual-py`](https://www.npmjs.com/package/contextual-py).
- **v1** - initial version of the tool. simple enough.

---

## made by

**[@anshdhariwal](https://github.com/anshdhariwal)** built this instead of actually the intended work. out of the box right?

---

*if this saved you time, drop a star or just silently appreciate it. both are cool.*
