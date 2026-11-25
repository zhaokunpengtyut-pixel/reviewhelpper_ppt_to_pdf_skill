---
name: courseware-review-pdf
description: Turn a folder of courseware slides (.ppt/.pptx) into a print-ready revision handout PDF — full-page screenshots of the original slides, every worked example re-explained step by step, tables and diagrams preferred over prose. Use when the user hands over lecture decks and asks to "make a revision PDF", "convert the courseware to PDF", "copy every question out with answers", "keep the original figures/screenshots", or "produce a handout I can print". Covers .ppt conversion, full slide export, question triage, HTML+CSS layout, Chrome PDF rendering with page numbers, and page-by-page QA.
whenToUse: The input is courseware slides (.ppt/.pptx) and the output must be a readable/printable revision handout, with the original figures and questions preserved.
---

# Courseware PPT → Revision PDF

This pipeline has been run end to end once: **8 `.ppt` decks / 679 slides → 108-page handout / 70 embedded slide screenshots / 72 examples fully solved / 15 MB.**
The `scripts/` and `assets/` in this directory are the verified versions — **copy them into the project and reuse them, do not rewrite them.**

## 0. Confirm three things before starting

1. **Shape of the deliverable**: page budget (10-page cheat sheet vs. 100-page full walkthrough), whether a cover / contents / appendix is wanted.
2. **Material boundary**: which decks, and whether an exercise file exists. (A common case: the `习题/` folder is empty — say so up front and offer a substitute.)
3. **Figure requirement**: the default is **embed the full slide as an image**. This is the heart of these requests — never paraphrase a circuit diagram.

If the user did not specify, apply this default contract and state it at delivery:

> ① **Every** exercise in the courseware is reproduced with a detailed, step-by-step answer; anything the courseware omits is written in.
> ② **If it can be a figure or a table, it must not be a paragraph** (truth tables, K-maps, state tables, comparison tables, flow diagrams).
> ③ **Embed the original slide screenshot**, captioned with its source page, followed by explanatory text.

## 1. Environment check (run it, do not assume)

```powershell
# Office / LibreOffice
Test-Path "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"
# Browser (for PDF rendering)
Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe"
# Python dependencies
& $PY -c "import pptx, PIL; print('ok')"
```

- With PowerPoint COM available, **prefer PowerPoint**: convert `.ppt` with `SaveAs(24)` for maximum fidelity. LibreOffice is the fallback only.
- No poppler / ghostscript / ImageMagick is fine — `scripts/pdf2png.ps1` (WinRT `Windows.Data.Pdf`) does the visual QA with zero dependencies.
- Use the bundled interpreter from `load_workspace_dependencies` for `$PY`. If `import pptx` fails, do not reflexively `pip install`; check the interpreter first.

## 2. The six-stage pipeline

### A. Convert + export every slide

```powershell
powershell -ExecutionPolicy Bypass -File scripts\ppt2png.ps1 -InputDir <courseware dir>
# -> <dir>/_pptx/chNN.pptx, chNN.name.txt
# -> <dir>/_slides/chNN/001.png ...   (1600x1200)
```

- Export **all** slides in one go. Do not export on demand: you will need them later for question triage, formula verification, and illustration picking.
- 600+ slides take roughly 5–8 minutes; run it in the background while doing stage B.
- `chNN` decouples slide folders from Chinese filenames (safer in scripts); the mapping lives in `chNN.name.txt`.

### B. Extract text, build the index

```powershell
& $PY scripts\extract.py --pptx <dir>\_pptx --out <dir>\_text
```

Produces `chNN.json`, `chNN.txt`, and `index.txt`. **The index is the foundation for everything that follows**: one line per slide with a `P` (has picture) / `T` (has table) flag and the first 110 characters.

**Read those UTF-8 files with the `read` tool — never through the console.** A Windows console is GBK and will turn Chinese into mojibake, wasting a whole round.

### C. Locate the questions

Sweep `index.txt` once (one `read` per chapter, with offset), then narrow down with regex:

```
例|例题|习题|练习|思考|作业        → strong signal
求|试画|画出|写出|化简|设计|分析|判断 → weak signal, need two or more
```

Three hard-won rules:

- **Worked examples spill over several slides** — a single problem routinely spans 2–8 pages (state diagram → K-map → drive equations → logic diagram). Record the whole page range, not just the title page.
- **Complement bars**: this deck used the `A'` convention, which survives text extraction intact. Decks that draw negations as separate shapes lose them completely — **for any formula-heavy slide, `read_image` the original before writing**.
- **End-of-chapter "习题 X.Y" often lists numbers only.** Do not guess textbook problems. Instead: list the numbers, then write an equivalent self-test with answers, and explain the substitution in the usage notes.

### D. Fix the structure

The skeleton that worked:

```
Cover (contained card, NOT full-bleed — see pitfalls)
Contents (chapters + example counts + per-chapter keywords + source page ranges)
Per chapter:
  ├ Chapter lead-in (where it sits in the course + the 2–3 places students lose marks)
  ├ Knowledge at a glance (tables only, plus 1–2 SVG flow diagrams)
  ├ Example cards ×N   ← the body
  │    ├ Header bar: 【Example 4-3】…  |  deck 4, P18–P20
  │    ├ Restated question (one sentence, for searchability)
  │    ├ Full-slide screenshot(s), 1–3 in solving order
  │    └ Analysis: idea → numbered steps → boxed conclusion → pitfall callout
  └ Chapter self-test (table: no. / question / answer)
Appendix: part cheat-sheet, top-20 pitfalls, comprehensive self-test
```

Budget roughly 1–1.5 pages per example; 72 examples plus 8 chapter summaries ≈ 100 pages. Use that as the planning baseline.

### E. Process images, write HTML

```powershell
& $PY scripts\prep_img.py --slides <dir>\_slides --out _src\img --sel ch03:031 ch03:038 ...
```

- The white-point stretch (`--thr 232`) flattens the deck's grid texture to pure white: ~400 KB → ~150 KB per slide, and it prints cleaner.
- 1200 px wide (≈235 dpi at 13 cm), quality 80. 70 images ≈ 11 MB, finished PDF ≈ 15 MB. Acceptable.
- Convert only the slides you will embed. `--all` only for very short decks or when the user explicitly wants every page.

Name chapter fragments `NN_*.html` (`00_front` … `09_appendix`) and pair them with `assets/print.css`. **Use that stylesheet as-is; do not start from scratch.**

### F. Render the PDF

```powershell
& $NODE scripts\print_pdf.mjs _src "..\xxx_revision-notes.pdf" "Digital Circuits · Revision Notes"
```

`print_pdf.mjs` does three things: concatenates the chapter fragments and inlines the CSS → launches Chrome over CDP → `Page.printToPDF`.
**It must go through CDP; do not use `chrome --headless --print-to-pdf`.** The CLI cannot customise the footer, and page numbers only exist in CDP's `footerTemplate`.

### G. QA — five checks, none optional

```powershell
& scripts\pdf2png.ps1 <pdf> _prev 200
```

1. **Page count**: does `pages: N` match the expected order of magnitude?
2. **Near-empty pages**: `Get-ChildItem _prev\*.png | ? Length -lt 160000`. Any hit means the layout broke (bad page break, missing image). **Target: zero.**
3. **Missing images**: regex every `src="..."` out of the HTML and `Test-Path` each one. **A missing image is a silent broken icon in the PDF — nobody notices until the user does.**
4. **Spot-check visually**: cover, contents, each chapter's first page, one SVG page, one code page.
5. **Spot-check content**: pick three examples at random and compare the formulas in the PDF against the original slide, character by character.

## 3. Layout rules (already implemented in `assets/print.css`)

| Item | Value | Why |
|---|---|---|
| Paper / margins | A4, `0.52 / 0.58 / 0.55 / 0.55` inch | Leaves room for the CDP footer |
| Body | Microsoft YaHei 10.2 pt, line-height 1.75 | Dense Chinese stays readable |
| Formulas | Cambria/Times italic, keep the courseware's `A'` notation | Matches the original figures; no mental translation for the reader |
| Chapter headings | `break-before: page` + dark blue bar | One chapter per page block, easy to flip |
| Tables | dark header, zebra rows, `break-inside: avoid` | A split table is the worst thing to read |
| Example cards | gradient header bar, light border, **allowed to split** | See pitfall 3 |
| Conclusion / pitfall | green `.res` box / red `.box.warn` box | Visual anchors for revision |

## 4. Content rules

- **Explain the *why*, not just the answer.** If the deck gives the answer, add the intermediate steps; if it gives only the conclusion, add the derivation; if it jumps, name the formula that was used.
- **Re-check the courseware's own answers.** The reference run found two real errors (one answer's third term was wrong and contradicted the truth table at a specific minterm; two example numbers were mislabelled). Flag them in the text as "the courseware prints X here; the K-map gives Y — trust the K-map." That is worth more than copying.
- **One notation throughout**: `A'` for NOT, `⊕` for XOR, `∑m(…)` for minterms, `∑d(…)` for don't-cares; subscripts via `<sub>`.
- **End every chapter with a self-test table** (no. / question / answer) so answers can be checked at a glance.
- **The appendix must fit on a few pages**: part cheat-sheet + top-20 pitfalls + comprehensive self-test. In the last 30 minutes before an exam, people only read the appendix.

## 5. Pitfalls (all of these actually happened)

1. **Do not full-bleed the cover.** Negative margins get clipped in print and a `position: fixed` footer does not sit reliably at the bottom of every page. Make the cover a **contained card inside the page**, and leave page numbers to the CDP `footerTemplate`.
2. **Page numbers only work via `footerTemplate`.** Chrome does not support CSS `@page { @bottom-center }`, and `position: fixed` drifts when the document paginates.
3. **`.card { break-inside: avoid }` manufactures near-empty pages** (121 pages, a dozen of them holding 3 lines). Correct approach: let cards split; give only the header bar `break-after: avoid`, and give tables / figures / `<pre>` `break-inside: avoid`. This alone took 121 pages down to 105.
4. **Justification stretches Chinese headings.** `body { text-align: justify }` reaches headings and produces `扩 展 题 三 步 速 算 法`. Set `text-align: left` on `h1, h2, h3, .box .h, .card > .hd`.
5. **PowerShell 5.1 reads `.ps1` as ANSI.** Chinese paths inside a script become "file not found". Keep scripts **ASCII-only** (pass paths as parameters), or save as UTF-8 **with BOM**.
6. **`$presentation.Visible = $true` throws an MsoTriState conversion error.** Use `$ppt.Visible = -1` inside try/catch.
7. **The console is GBK.** Write all Chinese analysis to files and `read` them back; never inspect Chinese through `python ... 2>&1 | Out-String`.
8. **`python-pptx` cannot read `.ppt`** (only `.pptx`). Convert first.
9. **`Slide.Export` dimensions**: 1600×1200 for 4:3 decks, 1600×900 for 16:9, otherwise images are stretched.
10. **Missing images fail silently.** A wrong `<img src>` leaves only a broken icon, easy to miss in a quick scroll — always verify with the script (QA check 3).
11. **Uncompressed images bloat the PDF.** 79 PNGs at 1400 px / JPEG-86 were 27 MB; 1200 px + white-point stretch + q80 is 11 MB.
12. **Do not mix full-width Chinese quotation marks into `<pre>` code blocks** — they misalign in monospace. Use Consolas, light grey background, blue left rule.

## 6. Delivery and cleanup

- Run all five QA checks before delivery, `present` the PDF once, and mention the source files in one sentence at the end.
- Intermediates are heavy (679 PNGs + pptx ≈ 230 MB). After delivery:
  - keep `_src/` (chapter HTML + images + CSS, ~11 MB) for rework;
  - keep `_build/` (`print_pdf.mjs` + `pdf2png.ps1`);
  - delete `_slides/`, `_pptx/`, `_text/`, `.chromeprofile/`, `_prev/`, and draft PDFs.
- Also delete **unreferenced** images from `_src/img/` (diff the HTML `src` set against the directory) — usually saves another 1–2 MB.

## 7. Reuse table

| Stage | Script / asset | Notes |
|---|---|---|
| Convert + export slides | `scripts/ppt2png.ps1` | Needs PowerPoint COM; set `-Width/-Height` to match the deck aspect ratio |
| Extract + index | `scripts/extract.py` | Recurses groups, reads notes, emits `index.txt` |
| Image compression | `scripts/prep_img.py` | `--sel ch03:031` or `--all` |
| Layout | `assets/print.css` | Copy to the project as `style.css` |
| Render PDF (with page numbers) | `scripts/print_pdf.mjs` | Chrome + CDP, `footerTemplate` for page numbers |
| Visual QA | `scripts/pdf2png.ps1` | WinRT, zero dependencies, replaces poppler |
