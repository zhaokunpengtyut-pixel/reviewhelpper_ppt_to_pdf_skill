# -*- coding: utf-8 -*-
"""Extract text / tables / picture counts from converted .pptx decks.

Usage:
    python extract.py --pptx _pptx --out _text

Outputs (all UTF-8):
    _text/ch01.json   per-slide {i, title, text, tables, pics, notes}
    _text/ch01.txt    human/agent readable dump  (=== SLIDE n ===)
    _text/index.txt   one compact line per slide: "P012 T <first 110 chars>"
                      P = has picture(s), T = has table

Design notes:
  * Recurses into group shapes (python-pptx does not flatten them).
  * Reads speaker notes if present.
  * Keep this script's output on disk and grep the FILES, never the console:
    a GBK console mangles Chinese and makes triage impossible.
"""
import argparse
import json
import os

from pptx import Presentation


def walk(shapes, acc, depth=0):
    for sh in shapes:
        try:
            if sh.shape_type == 6 or sh.__class__.__name__ == "GroupShape":
                walk(sh.shapes, acc, depth + 1)
                continue
        except Exception:
            pass

        try:
            if sh.has_text_frame and sh.text_frame.text.strip():
                acc["texts"].append(sh.text_frame.text)
        except Exception:
            pass

        try:
            if sh.has_table:
                acc["tables"].append([[c.text for c in r.cells] for r in sh.table.rows])
        except Exception:
            pass

        try:
            if sh.shape_type == 13:      # MSO_SHAPE_TYPE.PICTURE
                acc["pics"] += 1
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True, help="directory holding chNN.pptx")
    ap.add_argument("--out", required=True, help="directory for json/txt/index")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    index_lines = []

    for fn in sorted(f for f in os.listdir(args.pptx) if f.endswith(".pptx")):
        prs = Presentation(os.path.join(args.pptx, fn))
        tag = fn[:-5]
        rows = []
        for i, slide in enumerate(prs.slides, 1):
            acc = {"texts": [], "tables": [], "pics": 0}
            walk(slide.shapes, acc)
            notes = ""
            try:
                if slide.has_notes_slide:
                    notes = slide.notes_slide.notes_text_frame.text
            except Exception:
                pass
            title = ""
            try:
                if slide.shapes.title is not None:
                    title = slide.shapes.title.text
            except Exception:
                pass
            rows.append({
                "i": i,
                "title": title,
                "text": "\n".join(acc["texts"]),
                "tables": acc["tables"],
                "pics": acc["pics"],
                "notes": notes,
            })

        with open(os.path.join(args.out, tag + ".json"), "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)

        with open(os.path.join(args.out, tag + ".txt"), "w", encoding="utf-8") as f:
            for d in rows:
                f.write("=" * 70 + "\n")
                f.write("SLIDE %d   pics=%d  tables=%d\n" % (d["i"], d["pics"], len(d["tables"])))
                f.write("-" * 70 + "\n")
                f.write(d["text"] + "\n")
                for tb in d["tables"]:
                    f.write("[TABLE]\n")
                    for r in tb:
                        f.write(" | ".join(r) + "\n")
                if d["notes"].strip():
                    f.write("[NOTES] " + d["notes"] + "\n")

        index_lines.append("#### %s  (%d slides)" % (tag, len(rows)))
        for d in rows:
            firsts = [l.strip() for l in d["text"].split("\n") if l.strip()]
            head = " / ".join(firsts[:2])[:110]
            index_lines.append("%s%03d %s %s" % ("P" if d["pics"] else "-", d["i"],
                                                 "T" if d["tables"] else "-", head))
        print("%s -> %d slides" % (fn, len(rows)))

    with open(os.path.join(args.out, "index.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(index_lines) + "\n")
    print("index -> %s" % os.path.join(args.out, "index.txt"))


if __name__ == "__main__":
    main()
