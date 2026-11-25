# -*- coding: utf-8 -*-
"""Normalise selected slide PNGs into print-friendly JPEGs for the PDF.

Why normalise:
  * The PPT template usually carries a faint grid/paper texture. JPEG cannot
    compress it, so a 1600x1200 slide weighs ~400 KB. A white-point stretch
    (every channel value >= --thr becomes 255) flattens the background to pure
    white: text stays crisp, print looks cleaner, and each page drops to
    ~140-190 KB at 1200 px wide.
  * 1200 px across a ~13 cm-wide figure is ~235 dpi — plenty for print.

Usage:
    python prep_img.py --slides _slides --out _src/img --sel ch02:022 ch03:031 ch03:038
    python prep_img.py --slides _slides --out _src/img --all

--sel  takes <chapterTag>:<slideNumber> pairs (zero padding tolerated).
--all  converts every slide of every chapter (heavy; only when the deck is
       short or the user explicitly wants every page embedded).
"""
import argparse
import os

from PIL import Image


def flatten(im, thr):
    lut = [255 if v >= thr else int(v * 255 / thr) for v in range(256)]
    return im.point(lut * 3)


def convert(src, dst, width, quality, thr):
    im = Image.open(src).convert("RGB")
    if thr < 255:
        im = flatten(im, thr)
    h = int(im.height * width / im.width)
    im = im.resize((width, h), Image.LANCZOS)
    im.save(dst, "JPEG", quality=quality, optimize=True, progressive=True)
    return os.path.getsize(dst)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--slides", required=True, help="directory holding chNN/NNN.png")
    ap.add_argument("--out", required=True, help="output directory for JPEGs")
    ap.add_argument("--sel", nargs="*", default=[], help="ch02:022 ch03:031 ...")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--width", type=int, default=1200)
    ap.add_argument("--quality", type=int, default=80)
    ap.add_argument("--thr", type=int, default=232, help="white point, 255 disables")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    pairs = []

    if args.all:
        for ch in sorted(os.listdir(args.slides)):
            d = os.path.join(args.slides, ch)
            if not os.path.isdir(d):
                continue
            for f in sorted(os.listdir(d)):
                if f.lower().endswith(".png"):
                    pairs.append((ch, f[:-4]))
    else:
        for s in args.sel:
            if ":" not in s:
                raise SystemExit("bad --sel item %r, expected ch02:022" % s)
            ch, num = s.split(":", 1)
            pairs.append((ch, num.zfill(3)))

    total = 0
    missing = []
    for ch, num in pairs:
        src = os.path.join(args.slides, ch, num + ".png")
        if not os.path.exists(src):
            missing.append(src)
            continue
        dst = os.path.join(args.out, "%s_%s.jpg" % (ch, num))
        total += convert(src, dst, args.width, args.quality, args.thr)

    print("converted %d / %d  ->  %.1f MB" % (len(pairs) - len(missing), len(pairs), total / 1024 / 1024))
    if missing:
        print("MISSING (%d):" % len(missing))
        for m in missing:
            print("  " + m)


if __name__ == "__main__":
    main()
