# contextual 📄

> your documents walked into a tool. they left as one.

a dead-simple python CLI that grabs every **PDF**, **DOCX**, and **PPTX** in your current directory and smashes them into a single, clean `.txt` file. no config, no flags, no drama.

---

## why?

because copy-pasting 37 files into one doc at 2am for a context file is **not it**.

contextual does the boring part so you can focus on pretending you read all of them.

---

## how it works

```
you run it → it finds your files → you press enter → you name the output → done.
```

that's it. that's the tool.

---

## quick start

```bash
cd whatever-folder
python contextual.py
```

---

## requirements

- python 3.10+
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

basically it looks like someone organized your stuff. finally.

## use cases

- **sending context to AI** - merge all your project docs into one file and feed it to ChatGPT, Claude, Gemini, or whatever LLM you're prompting. one file, full context, no copy-paste marathon.
- **studying for exams** - dump all your lecture slides and notes into a single searchable text file. ctrl+F is now your best friend.
- **project documentation** - combine scattered reports, specs, and presentations into one reference doc before a meeting or deadline.
- **code review prep** - merge requirement docs, design specs, and test plans into a single file so reviewers have everything in one place.
- **archiving** - turn a folder of mixed-format documents into a single plain text backup that'll open on literally anything, forever.

---

## faq- you may ask

**Q: does it support xlsx?**
A: no. this is contextual, not excel anonymous.

**Q: can i reorder the files?**
A: they sort by leading number automatically. name them `1-intro.pdf`, `2-body.docx`, etc. or don't. i'm not your mom.

**Q: it says "making your file buddy" and waits 2 seconds even though it's already done.**
A: vibes. it's called pacing.

**Q: why doesn't page numbering work for .docx files?**
A: because `.docx` files don't actually have "pages" as they're flow documents. page breaks only exist when Microsoft Word renders them on screen, and that info isn't stored in the file itself. so there's nothing to number. PDFs have real pages, PPTXs have real slides, DOCX has paragraphs. we just dump them all.

---

## changelog

- **v1** - initial version of the tool. simple enough.

---

## made by

**[@anshdhariwal](https://github.com/anshdhariwal)** built this instead of actually reading the documents.

---

*if this saved you time, drop a star or just silently appreciate it. both are cool.*
