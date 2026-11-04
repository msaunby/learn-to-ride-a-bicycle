# /// script
# requires-python = ">=3.11"
# dependencies = ["python-pptx>=1.0", "pillow"]
# ///
"""Build editable PowerPoint decks from the slide markdown in slides/marp/.

    uv run slides/build_pptx.py                      # every deck
    uv run slides/build_pptx.py slides/marp/01-*.md  # selected decks

Every slide is drawn with native PowerPoint text boxes, shapes and tables, so
the output can be edited in PowerPoint. The markdown uses the same slide
classes as the Marp theme, so the same file also renders with Marp for a
quick PDF preview. Colours, fonts and course details come from course.toml;
each run also writes the Marp theme, marp/<theme name>.css, from them and
theme-base.css.

Supported markdown: # / ## / ### headings, paragraphs, - and 1. lists (an
indented second line is the item's description), > blockquotes, pipe tables,
**bold**, *italic*, `code`, a paragraph wrapped in <mark> (lead panel),
![alt](path) images on a line of their own (path relative to the markdown
file; Marp size keywords such as h:400 in the alt text are ignored), and
<!-- comments --> which become speaker notes.
"""

import math
import re
import sys
import tomllib
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

from lxml import etree
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = Path(__file__).resolve().parent

# Slide is laid out in CSS pixels (1280 x 720, the Marp size) and converted.
W, H = 1280, 720
MARGIN = 64
CONTENT_W = W - 2 * MARGIN

CONFIG = tomllib.loads((HERE.parent / "course.toml").read_text(encoding="utf-8"))
THEME = CONFIG["theme"]
COLOURS = THEME["colours"]
FONTS = THEME["fonts"]

PRIMARY = COLOURS["primary"]
SECONDARY = COLOURS["secondary"]
ACCENT = COLOURS["accent"]
BACKGROUND = COLOURS["background"]
WHITE = COLOURS["white"]
TEXT = COLOURS["text"]
MUTED = COLOURS["muted"]
CALLOUT = COLOURS["callout"]
ACCENT_LIGHT = COLOURS["accent_light"]
POSITIVE = COLOURS["positive"]
NEGATIVE = COLOURS["negative"]
BORDER = COLOURS["border"]
STRIPE = COLOURS["stripe"]
ON_DARK_MUTED = COLOURS["on_dark_muted"]  # secondary text on dark slides

HEADING = FONTS["heading"]
BODY = FONTS["body"]


def px(v):
    return Emu(int(round(v * 9525)))


# ---------------------------------------------------------------- parsing


@dataclass
class Slide:
    classes: set
    blocks: list
    notes: list = field(default_factory=list)


ITEM = re.compile(r"(-|\*|\d+\.)\s+(.*)")
IMAGE = re.compile(r"!\[(.*?)\]\((.+?)\)\s*$")


class Items(list):
    """List items; `start` is the number of the first item of an ordered list."""
    start = 1


def parse_deck(text):
    meta = {}
    m = re.match(r"---\n(.*?)\n---\n", text, re.S)
    if m:
        for line in m.group(1).splitlines():
            if ":" in line:
                k, v = line.split(":", 1)
                meta[k.strip()] = v.strip()
        text = text[m.end():]
    parts = re.split(r"^---\s*$", text, flags=re.M)
    return meta, [parse_slide(p) for p in parts if p.strip()]


def parse_slide(src):
    classes, notes = set(), []

    def comment(m):
        body = m.group(1).strip()
        d = re.match(r"_?class:\s*(.+)", body)
        if d:
            classes.update(d.group(1).split())
        elif not re.match(r"_?\w+:", body):
            notes.append(body)
        return ""

    src = re.sub(r"<!--(.*?)-->", comment, src, flags=re.S)
    lines = src.split("\n")
    blocks, i = [], 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
        elif m := re.match(r"(#{1,3})\s+(.*)", line):
            blocks.append((f"h{len(m.group(1))}", m.group(2).strip()))
            i += 1
        elif m := IMAGE.match(line.strip()):
            alt = " ".join(w for w in m.group(1).split() if not re.match(r"\w+:", w))
            blocks.append(("image", (alt, m.group(2))))
            i += 1
        elif line.startswith("```"):
            code, i = [], i + 1
            while i < len(lines) and not lines[i].startswith("```"):
                code.append(lines[i].rstrip())
                i += 1
            blocks.append(("code", code))
            i += 1
        elif line.startswith(">"):
            quote = []
            while i < len(lines) and lines[i].startswith(">"):
                if t := lines[i][1:].strip():
                    quote.append(t)
                i += 1
            blocks.append(("quote", quote))
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-+:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            blocks.append(("table", rows))
        elif m := ITEM.match(line):
            kind = "ol" if m.group(1)[0].isdigit() else "ul"
            items = Items()
            items.start = int(m.group(1)[:-1]) if kind == "ol" else 1
            while i < len(lines):
                if m := ITEM.match(lines[i]):
                    items.append([m.group(2).strip()])
                elif items and lines[i][:1] in (" ", "\t") and lines[i].strip():
                    items[-1].append(lines[i].strip())
                else:
                    break
                i += 1
            blocks.append((kind, items))
        else:
            para = [line.strip()]
            i += 1
            while i < len(lines) and lines[i].strip() and not re.match(r"#|>|\||```|[-*]\s|\d+\.\s", lines[i]):
                para.append(lines[i].strip())
                i += 1
            text = " ".join(para)
            if text.startswith("<mark>"):
                blocks.append(("lead", re.sub(r"</?mark>", "", text)))
            else:
                blocks.append(("p", text))
    return Slide(classes, blocks, notes)


def runs(text):
    """Split inline markdown into (text, kind) runs; kind is '', 'b', 'i' or 'c'."""
    text = re.sub(r"<br\s*/?>", "\n", text)
    out = []
    for tok in re.split(r"(\*\*.+?\*\*|\*[^*\s][^*]*?\*|`[^`]+`)", text):
        if not tok:
            continue
        if tok.startswith("**"):
            out.append((tok[2:-2], "b"))
        elif tok.startswith("`"):
            out.append((tok[1:-1], "c"))
        elif tok.startswith("*") and len(tok) > 2:
            out.append((tok[1:-1], "i"))
        else:
            out.append((tok, ""))
    return out


def plain(text):
    return "".join(t for t, _ in runs(text))


def title_desc(item):
    """A list item with an indented second line is a titled card/row."""
    if len(item) > 1:
        return plain(item[0]), " ".join(item[1:])
    return None, item[0]


FONT_FILES = {
    (BODY, False): FONTS["body_files"][0], (BODY, True): FONTS["body_files"][1],
    (HEADING, False): FONTS["heading_files"][0], (HEADING, True): FONTS["heading_files"][1],
}
FONT_DIRS = [
    Path("/mnt/c/Windows/Fonts"), Path("C:/Windows/Fonts"),
    Path("/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
    Path.home() / "Library/Fonts", Path.home() / ".fonts", Path("/usr/share/fonts"),
]


@lru_cache(maxsize=None)
def measurer(font, bold, size):
    """A function giving the width of a string in the real font, if installed."""
    name = FONT_FILES[(font, bold)]
    for d in FONT_DIRS:
        for path in [d / name, d / name.capitalize()] + (list(d.rglob(name)) if d.name == "fonts" and d.is_dir() else []):
            if path.is_file():
                return ImageFont.truetype(str(path), size).getlength
    return None


def text_h(text, width, size, font=BODY, line=1.1, bold=None):
    """Rendered height of wrapped text, used to stack shapes.

    Wraps word by word with the real font metrics when the course fonts are
    installed, otherwise falls back to an average character width.
    """
    bold = font == HEADING if bold is None else bold
    measure = measurer(font, bold, size)
    n = 0
    for para in plain(text or "").split("\n"):
        if measure is None:
            char_w = size * (0.56 if font == HEADING else 0.5)
            n += max(1, math.ceil(len(para) * char_w / width))
            continue
        lines, current = 1, ""
        for word in para.split():
            trial = f"{current} {word}".strip()
            if current and measure(trial) > width:
                lines, current = lines + 1, word
            else:
                current = trial
        n += lines
    return max(n, 1) * size * 1.22 * line


def text_w(text, size, bold=False):
    measure = measurer(BODY, bold, size)
    return measure(text) if measure else len(text) * size * 0.5


def column_widths(rows, w, size, pad):
    """Fit columns to content, never narrower than a column's longest word."""
    ncols = len(rows[0])
    bold = lambda r, cell: r == 0 or cell.startswith("**")
    least = [max(text_w(word, size, bold(r, row[c])) for r, row in enumerate(rows)
                 for word in plain(row[c]).split() or [""]) + pad + 6 for c in range(ncols)]
    full = [max(text_w(plain(row[c]), size, bold(r, row[c])) for r, row in enumerate(rows)) + pad + 6
            for c in range(ncols)]
    if sum(full) <= w:
        return [f * w / sum(full) for f in full]
    spare = [f - m for f, m in zip(full, least)]
    return [m + (w - sum(least)) * s / sum(spare) for m, s in zip(least, spare)]


# ---------------------------------------------------------------- drawing


def set_alpha(fill_parent, alpha):
    srgb = fill_parent.find(".//" + qn("a:srgbClr"))
    etree.SubElement(srgb, qn("a:alpha")).set("val", str(int(alpha * 100000)))


def box(slide, x, y, w, h, fill=None, line=None, radius=0, shape=None, alpha=None):
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    s = slide.shapes.add_shape(kind, px(x), px(y), px(w), px(h))
    if radius:
        s.adjustments[0] = radius / min(w, h)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(fill)
        if alpha is not None:
            set_alpha(s.fill._xPr, alpha)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = RGBColor.from_string(line)
        s.line.width = px(1)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def fill_para(p, text, size=26, color=TEXT, font=BODY, bold=False, italic=False,
              strong=None, align=None, line=1.1, before=0, raw=False):
    p.line_spacing = line
    if before:
        p.space_before = Pt(before * 0.75)
    if align:
        p.alignment = align
    for t, kind in [(text, "")] if raw else runs(text):
        r = p.add_run()
        r.text = t
        f = r.font
        f.name = "Consolas" if kind == "c" else font
        f.size = Pt(size * 0.75)
        f.bold = bold or kind == "b"
        f.italic = italic or kind == "i"
        f.color.rgb = RGBColor.from_string(strong if kind == "b" and strong else color)


def text_box(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, **style):
    """paras: list of markdown strings, or (string, style-overrides) pairs."""
    tb = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for k, para in enumerate(paras):
        text, extra = para if isinstance(para, tuple) else (para, {})
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        fill_para(p, text, **{**style, **extra})
    return tb


def bullet(p, char, color, indent):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(px(indent)))
    pPr.set("indent", str(-px(indent)))
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)


def badge(slide, x, y, d, fill, color, label, size=18):
    s = box(slide, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    fill_para(tf.paragraphs[0], str(label), size=size, color=color, bold=True, align=PP_ALIGN.CENTER)
    return s


# ---------------------------------------------------------------- blocks


class Deck:
    def __init__(self, slide, classes):
        self.s = slide
        self.cls = classes
        self.dark = bool(classes & {"title", "dark", "takeaway"})

    # Plain blocks -------------------------------------------------------

    def p(self, text, x, y, w):
        h = text_h(text, w, 26)
        text_box(self.s, x, y, w, h, [text], size=26,
                 color=WHITE if self.dark else TEXT, strong=WHITE if self.dark else PRIMARY)
        return y + h

    def lead(self, text, x, y, w):
        h = text_h(text, w - 44, 22) + 28
        box(self.s, x, y, w, h, fill=ACCENT_LIGHT, radius=10)
        text_box(self.s, x + 22, y + 14, w - 44, h - 28, [text], size=22,
                 color=TEXT, strong=PRIMARY, anchor=MSO_ANCHOR.MIDDLE)
        return y + h

    def heading(self, text, x, y, w, size, color):
        h = text_h(text, w, size, HEADING)
        text_box(self.s, x, y, w, h, [text], size=size, color=color, font=HEADING, bold=True)
        return y + h

    def h2(self, text, x, y, w):
        return self.heading(text, x, y, w, 30, ACCENT if self.dark else PRIMARY)

    def h3(self, text, x, y, w):
        return self.heading(text, x, y, w, 27, ACCENT if self.dark else SECONDARY)

    def quote(self, lines, x, y, w):
        pad_x, pad_y = 26, 20
        inner = w - 2 * pad_x
        paras, h = [], 0
        for k, line in enumerate(lines):
            label = re.fullmatch(r"\*\*([^*]+)\*\*", line)
            before = 10 if k else 0
            if label:
                paras.append((label.group(1), dict(size=18, bold=True, color=ACCENT if self.dark else SECONDARY, before=before)))
                h += text_h(line, inner, 18) + before
            else:
                size = 26 if self.dark else 22
                paras.append((line, dict(size=size, before=before)))
                h += text_h(line, inner, size) + before
        h += 2 * pad_y
        if self.dark:
            box(self.s, x, y, w, h, fill=WHITE, alpha=0.07, radius=10)
        else:
            box(self.s, x, y, w, h, fill=CALLOUT, radius=10)
        text_box(self.s, x + pad_x, y + pad_y, inner, h - 2 * pad_y, paras,
                 color=WHITE if self.dark else PRIMARY, strong=ACCENT if self.dark else SECONDARY)
        return y + h

    def ul(self, items, x, y, w):
        if "cards" in self.cls:
            return self.cards(items, x, y, w)
        size = 26
        tb = text_box(self.s, x, y, w, 10, [])
        tf = tb.text_frame
        h = 0
        for k, item in enumerate(items):
            text = "<br>".join(item)
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            fill_para(p, text, size=size, color=WHITE if self.dark else TEXT,
                      strong=WHITE if self.dark else PRIMARY, before=10 if k else 0)
            bullet(p, "■", ACCENT, 28)
            h += text_h(text, w - 28, size) + (10 if k else 0)
        tb.height = px(h)
        return y + h

    def ol(self, items, x, y, w):
        if "takeaway" in self.cls:
            return self.takeaway_rows(items, x, y, w)
        if "dark" in self.cls:
            return self.dark_rows(items, x, y, w)
        if "teal" in self.cls:
            return self.boxed_rows(items, x, y, w)
        return self.steps(items, x, y, w)

    def table(self, rows, x, y, w):
        ncols = len(rows[0])
        widths = column_widths(rows, w, 23, 36)
        heights = [max(46, max(text_h(r[c], widths[c] - 36, 23) for c in range(ncols)) + 20) for r in rows]
        gt = self.s.shapes.add_table(len(rows), ncols, px(x), px(y), px(w), px(sum(heights)))
        tbl = gt.table
        # "No Style, No Grid": all colours come from the cell fills below
        gt._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId")).text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
        tbl.first_row = True
        tbl.horz_banding = False
        for c, cw in enumerate(widths):
            tbl.columns[c].width = px(cw)
        for r, row in enumerate(rows):
            tbl.rows[r].height = px(heights[r])
            for c, value in enumerate(row):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = px(18)
                cell.margin_top = cell.margin_bottom = px(4)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(
                    SECONDARY if r == 0 else STRIPE if r % 2 == 0 else WHITE)
                tf = cell.text_frame
                tf.word_wrap = True
                fill_para(tf.paragraphs[0], value, size=23, line=1.0,
                          color=WHITE if r == 0 else TEXT, bold=r == 0, strong=PRIMARY if r else WHITE)
                if r:
                    ln = etree.Element(qn("a:lnB"), w="9525")
                    solid = etree.SubElement(ln, qn("a:solidFill"))
                    etree.SubElement(solid, qn("a:srgbClr")).set("val", BORDER)
                    cell._tc.get_or_add_tcPr().insert(0, ln)
        return y + sum(heights)

    def code(self, lines, x, y, w):
        size, pad = 20, 20
        h = len(lines) * size * 1.22 + 2 * pad
        box(self.s, x, y, w, h, fill=PRIMARY if not self.dark else COLOURS["code_dark"], radius=8)
        text_box(self.s, x + 22, y + pad, w - 44, h - 2 * pad, lines, size=size, color=WHITE,
                 font="Consolas", line=1.0, raw=True)
        return y + h

    def image(self, data, x, y, w):
        """Scale to fit the width and the space left on the slide, centred."""
        alt, path = data
        with Image.open(path) as im:
            iw, ih = im.size
        scale = min(w / iw, (H - 44 - y) / ih)
        pw, ph = iw * scale, ih * scale
        pic = self.s.shapes.add_picture(str(path), px(x + (w - pw) / 2), px(y), px(pw), px(ph))
        pic._element.nvPicPr.cNvPr.set("descr", alt)
        return y + ph

    # Slide-class layouts ------------------------------------------------

    def cards(self, items, x, y, w):
        cols = 3 if "cols-3" in self.cls else 2
        numbered = "plain" not in self.cls
        gap, pad = 26, 28
        cw = (w - gap * (cols - 1)) / cols
        tx = pad + 64 if numbered else pad
        tw = cw - tx - pad
        bars = [SECONDARY, ACCENT, PRIMARY]
        for start in range(0, len(items), cols):
            row = [title_desc(it) for it in items[start:start + cols]]
            hgt = max(2 * pad + (text_h(t, tw, 27, HEADING) + 8 if t else 0) + text_h(d, tw, 21) for t, d in row)
            for c, (title, desc) in enumerate(row):
                n = start + c
                cx = x + c * (cw + gap)
                if numbered:
                    box(self.s, cx, y, cw, hgt, fill=WHITE, line=BORDER, radius=10)
                    badge(self.s, cx + 24, y + pad - 2, 42, CALLOUT, SECONDARY, n + 1)
                else:
                    box(self.s, cx, y, cw, hgt, fill=WHITE, line=BORDER)
                    box(self.s, cx, y, cw, 5, fill=bars[n % 3])
                ty = y + pad
                if title:
                    th = text_h(title, tw, 27, HEADING)
                    colour = ACCENT if not numbered and n % 3 == 1 else PRIMARY
                    text_box(self.s, cx + tx, ty, tw, th, [title], size=27, font=HEADING, bold=True, color=colour)
                    ty += th + 8
                text_box(self.s, cx + tx, ty, tw, hgt - (ty - y) - pad, [desc], size=21, color=MUTED, strong=PRIMARY)
            y += hgt + gap
        return y - gap

    def steps(self, items, x, y, w):
        tw = w - 58
        for n, item in enumerate(items, items.start):
            title, desc = title_desc(item)
            badge(self.s, x, y, 38, ACCENT, WHITE, n)
            th = text_h(title, tw, 25, HEADING) if title else 0
            if title:
                text_box(self.s, x + 58, y, tw, th, [title], size=25, font=HEADING, bold=True, color=PRIMARY)
            size = 21 if title else 25
            dh = text_h(desc, tw, size)
            text_box(self.s, x + 58, y + th, tw, dh, [desc], size=size,
                     color=MUTED if title else TEXT, strong=PRIMARY)
            y += max(38, th + dh) + 14
        return y - 14

    def boxed_rows(self, items, x, y, w):
        tx, tw = 78, w - 78 - 22
        for n, item in enumerate(items, items.start):
            title, desc = title_desc(item)
            th = text_h(title, tw, 25, HEADING) if title else 0
            dh = text_h(desc, tw, 21)
            h = th + dh + 26
            box(self.s, x, y, w, h, fill=WHITE, line=BORDER, radius=10)
            badge(self.s, x + 18, y + (h - 38) / 2, 38, PRIMARY, WHITE, n)
            if title:
                text_box(self.s, x + tx, y + 13, tw, th, [title], size=25, font=HEADING, bold=True, color=PRIMARY)
            text_box(self.s, x + tx, y + 13 + th, tw, dh, [desc], size=21, color=MUTED, strong=PRIMARY)
            y += h + 14
        return y - 14

    def dark_rows(self, items, x, y, w):
        tx, tw = 78, w - 78 - 22
        for n, item in enumerate(items, items.start):
            text = " ".join(item)
            h = text_h(text, tw, 25) + 36
            box(self.s, x, y, w, h, fill=WHITE, alpha=0.07, radius=10)
            badge(self.s, x + 20, y + (h - 38) / 2, 38, ACCENT, WHITE, n)
            text_box(self.s, x + tx, y, tw, h, [text], size=25, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
            y += h + 12
        return y - 12

    def takeaway_rows(self, items, x, y, w):
        tx, tw = 88, w - 88 - 26
        for n, item in enumerate(items, items.start):
            title, desc = title_desc(item)
            title = title or plain(desc)
            desc = desc if len(item) > 1 else ""
            th = text_h(title, tw, 27, bold=True)
            dh = text_h(desc, tw, 21) + 4 if desc else 0
            h = th + dh + 44
            box(self.s, x, y, w, h, fill=WHITE, alpha=0.07, radius=10)
            badge(self.s, x + 24, y + (h - 38) / 2, 38, ACCENT, WHITE, n)
            paras = [(title, dict(size=27, bold=True, color=WHITE))]
            if desc:
                paras.append((desc, dict(size=21, color=ON_DARK_MUTED, before=4)))
            text_box(self.s, x + tx, y, tw, h, paras, anchor=MSO_ANCHOR.MIDDLE, strong=WHITE)
            y += h + 18
        return y - 18

    def compare(self, blocks, x, y, w):
        """Columns from "### heading" + list pairs; other blocks follow below."""
        columns, rest, i = [], [], 0
        while i < len(blocks):
            if blocks[i][0] == "h3" and i + 1 < len(blocks) and blocks[i + 1][0] == "ul":
                columns.append((blocks[i][1], blocks[i + 1][1]))
                i += 2
            else:
                rest.append(blocks[i])
                i += 1
        gap, pad = 28, 26
        cw = (w - gap * (len(columns) - 1)) / len(columns)
        inner = cw - 2 * pad - 30
        heights = [24 + text_h(head, cw, 28, HEADING) + 12
                   + sum(text_h(" ".join(it), inner, 23) + 14 for it in items) + 12
                   for head, items in columns]
        hgt = max(heights)
        for c, (head, items) in enumerate(columns):
            good = c == 0
            colour = POSITIVE if good else NEGATIVE
            cx = x + c * (cw + gap)
            box(self.s, cx, y, cw, hgt, fill=WHITE, line=BORDER)
            box(self.s, cx, y, cw, 5, fill=colour)
            hh = text_h(head, cw, 28, HEADING)
            text_box(self.s, cx + pad, y + 22, cw - 2 * pad, hh, [head], size=28, font=HEADING, bold=True, color=colour)
            iy = y + 22 + hh + 12
            for it in items:
                text = " ".join(it)
                ih = text_h(text, inner, 23)
                text_box(self.s, cx + pad, iy, 30, ih, ["✓" if good else "✗"], size=23, bold=True, color=colour)
                text_box(self.s, cx + pad + 30, iy, inner, ih, [text], size=23, color=TEXT, strong=PRIMARY)
                iy += ih + 14
        if rest:
            return self.flow(rest, x, y + hgt + 22, w)
        return y + hgt

    def exercises(self, blocks, x, y, w):
        i = 0
        while i < len(blocks):
            kind, text = blocks[i]
            if kind != "h3":
                y = self.block(blocks[i], x, y, w) + 20
                i += 1
                continue
            desc = blocks[i + 1][1] if i + 1 < len(blocks) and blocks[i + 1][0] == "p" else ""
            i += 2 if desc else 1
            m = re.match(r"(.*?)\s*\*\(?([^*]+?)\)?\*\s*$", text)
            title, pill = (m.group(1), m.group(2)) if m else (text, None)
            pad = 28
            tw = w - 2 * pad - 6 - (110 if pill else 0)
            th = text_h(title, tw, 27, HEADING)
            dh = text_h(desc, w - 2 * pad - 6, 22) if desc else 0
            h = 20 + th + (6 + dh if desc else 0) + 20
            box(self.s, x, y, w, h, fill=WHITE, line=BORDER, radius=8)
            box(self.s, x, y, 6, h, fill=ACCENT)
            text_box(self.s, x + 6 + pad, y + 20, tw, th, [title], size=27, font=HEADING, bold=True, color=PRIMARY)
            if pill:
                p = box(self.s, x + w - pad - 90, y + 22, 90, 28, fill=ACCENT, radius=14)
                tf = p.text_frame
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                fill_para(tf.paragraphs[0], pill, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            if desc:
                text_box(self.s, x + 6 + pad, y + 26 + th, w - 2 * pad - 6, dh, [desc], size=22, color=MUTED, strong=PRIMARY)
            y += h + 14
        return y - 14

    def block(self, block, x, y, w):
        kind, data = block
        return getattr(self, kind)(data, x, y, w)

    def flow(self, blocks, x, y, w, gap=22):
        for b in blocks:
            y = self.block(b, x, y, w) + gap
        return y - gap


# ---------------------------------------------------------------- slides


def background(slide, colour):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(colour)


def circles(slide):
    box(slide, W - 300, -300, 600, 600, fill=WHITE, alpha=0.06, shape=MSO_SHAPE.OVAL)
    box(slide, W * 0.92 - 170, H * 0.18 - 170, 340, 340, fill=WHITE, alpha=0.04, shape=MSO_SHAPE.OVAL)
    box(slide, -210, H - 210, 420, 420, fill=SECONDARY, alpha=0.55, shape=MSO_SHAPE.OVAL)


def title_slide(slide, blocks):
    background(slide, PRIMARY)
    circles(slide)
    get = lambda kind: [d for k, d in blocks if k == kind]
    label, title, paras = get("h2"), get("h1"), get("p")
    x, w = 96, 860
    th = text_h(title[0], w, 66, HEADING, line=1.0) if title else 0
    total = 40 + th + 32 + 44 * len(paras)
    y = (H - total) / 2
    if label:
        text_box(slide, x, y, w, 24, [label[0].upper()], size=18, bold=True, color=ACCENT)
    y += 40
    if title:
        text_box(slide, x, y, w, th, [title[0]], size=66, font=HEADING, color=WHITE, line=1.0)
    y += th + 32
    for k, para in enumerate(paras):
        last = k == len(paras) - 1 and len(paras) > 1
        if last:
            y += 30
        text_box(slide, x, y, w, 40, [para], size=18 if last else 30, color=ACCENT if last else WHITE)
        y += 44


def render(prs, sl, number, name):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if sl.notes:
        slide.notes_slide.notes_text_frame.text = "\n\n".join(sl.notes)
    if "title" in sl.classes:
        title_slide(slide, sl.blocks)
        return
    deck = Deck(slide, sl.classes)
    background(slide, PRIMARY if deck.dark else BACKGROUND)
    if "takeaway" in sl.classes:
        circles(slide)

    blocks = list(sl.blocks)
    y = 44
    if blocks and blocks[0][0] == "h1":
        text_box(slide, MARGIN, y, CONTENT_W, 60, [blocks.pop(0)[1]], size=44, font=HEADING,
                 color=WHITE if deck.dark else PRIMARY)
        y += 60 + 26

    if "compare" in sl.classes:
        end = deck.compare(blocks, MARGIN, y, CONTENT_W)
    elif "exercises" in sl.classes:
        end = deck.exercises(blocks, MARGIN, y, CONTENT_W)
    elif "steps" in sl.classes and any(k == "quote" for k, _ in blocks) and not deck.dark:
        side_w, gap = 400, 40
        main = [b for b in blocks if b[0] != "quote"]
        side = [b for b in blocks if b[0] == "quote"]
        end = max(deck.flow(main, MARGIN, y, CONTENT_W - side_w - gap),
                  deck.flow(side, W - MARGIN - side_w, y, side_w))
    else:
        end = deck.flow(blocks, MARGIN, y, CONTENT_W)

    if end > H - 36:
        print(f"warning: {name} slide {number} may overflow ({end:.0f}px of {H}px)", file=sys.stderr)

    text_box(slide, W - 32 - 60, H - 36, 60, 20, [str(number)], size=13,
             color=COLOURS["page_number_dark"] if deck.dark else COLOURS["page_number"], align=PP_ALIGN.RIGHT)


def build(md_path, out_path):
    meta, slides = parse_deck(md_path.read_text(encoding="utf-8"))
    for sl in slides:
        sl.blocks = [(k, (d[0], md_path.parent / d[1])) if k == "image" else (k, d) for k, d in sl.blocks]
    prs = Presentation()
    prs.slide_width, prs.slide_height = px(W), px(H)
    props = prs.core_properties
    props.title = meta.get("title", md_path.stem)
    props.author = props.last_modified_by = CONFIG["course"]["author"]
    props.comments = props.subject = props.keywords = ""
    for n, sl in enumerate(slides, 1):
        render(prs, sl, n, md_path.name)
    prs.save(out_path)
    print(f"{md_path.name} => {out_path}")


def write_theme():
    """Write the Marp theme: the colours and fonts in course.toml + theme-base.css."""
    root = "".join(f"  --{k.replace('_', '-')}: #{v};\n" for k, v in COLOURS.items())
    css = f"""/* @theme {THEME['name']} */
/* @size 16:9 1280px 720px */

/* Generated by slides/build_pptx.py from course.toml and slides/theme-base.css.
   Don't edit: change those files and run slides/marp/build.sh. */

@import url('{FONTS['web_fonts']}');

:root {{
{root}  --heading-font: '{HEADING}', {FONTS['heading_fallback']};
  --body-font: '{BODY}', {FONTS['body_fallback']};
  --shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
}}

""" + (HERE / "theme-base.css").read_text(encoding="utf-8")
    path = HERE / "marp" / f"{THEME['name']}.css"
    if not path.is_file() or path.read_text(encoding="utf-8") != css:
        path.write_text(css, encoding="utf-8")
        print(f"theme => {path}")


def main():
    write_theme()
    paths = [Path(p) for p in sys.argv[1:]] or sorted((HERE / "marp").glob("[0-9]*.md"))
    for md in paths:
        build(md, HERE / f"{md.stem}.pptx")


if __name__ == "__main__":
    main()
