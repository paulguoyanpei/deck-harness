"""Find layout defects in a .pptx from shape geometry and font metrics.

Reports occlusion, container escape, text overflow, alignment defects, and
severe content imbalance without rendering the deck, so it can be run while
the deck is still being built. It is the substitute for visual QA when image
inspection is unavailable.

Usage:
    python layout_check.py deck.pptx [--waivers deck.layout.json] [--font-dir DIR]

Examples:
    python layout_check.py deck.pptx
    python layout_check.py deck.pptx --waivers deck.layout.json

Element naming
    Elements are expected to carry a path name (pptxgenjs `objectName`):

        objectName: "card1"        a container
        objectName: "card1/title"  sits on card1; may overlap it
        objectName: "@bg/blob"     background art; anything may cover it

    Elements that share no ancestor/descendant relationship must not overlap.
    Unnamed elements fall back to geometric inference, which is noisier and
    cannot be waived.

Waivers (deck.layout.json)
    {"waivers": [{"slide": 5, "a": "chart/bar3", "b": "chart/callout",
                  "ratio": 0.22, "reason": "callout deliberately marks the bar"}]}

    A waiver holds while the actual overlap stays within `ratio` + 0.10.
    Overlaps at or above 60% are never waivable.
"""

import argparse
import json
import math
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Emu

try:
    from fontTools.ttLib import TTCollection, TTFont
except ImportError:  # pragma: no cover - fontTools is a declared dependency
    TTFont = TTCollection = None


TOL = 0.02              # inches; ignore hairline nudges and rounding
OVERLAP_MIN = 0.03      # inches; below this an intersection is not visible
CENTER_TOL = 0.01       # inches; concentricity slack for centred text
BASELINE_TOL = 0.05     # inches; sibling text centreline slack
CONNECTOR_TOL = 0.06    # inches; arrow endpoint may sit just outside a node
SEVERE_RATIO = 0.60     # at or above this an overlap is never waivable
WAIVER_SLACK = 0.10     # a waiver survives this much extra coverage
BG_PREFIX = "@bg"
DEFAULT_FONT_SIZE = 18.0
LINE_HEIGHT = 1.2       # multiple of font size
LATIN_EM_RATIO = 0.6    # fallback advance width when metrics are unavailable
INFERRED_BG_ALPHA = 85  # unnamed fills below this opacity are treated as art
CONTENT_GRID_COLUMNS = 120
CONTENT_GRID_ROWS = 72
CONTENT_BODY_BOUNDS = (0.05, 0.16, 0.95, 0.88)
CONTENT_MIN_TEXT_CHARS = 60
CONTENT_EMPTY_HALF_PERCENT = 5.0
CONTENT_OCCUPIED_HALF_PERCENT = 25.0
CONTENT_EMPTY_BAND_RATIO = 0.25

AUTO_NAME_RE = re.compile(r"^(?:Shape|Text|Image|Chart|Table|Media|Picture)\s+\d+$")
DEFAULT_FONT_DIRS = ("/mnt/c/Windows/Fonts", "/usr/share/fonts", "/Library/Fonts")

# Filename stems for fonts whose family name does not match their file name.
FONT_ALIASES = {
    "microsoft yahei": ("msyh",),
    "microsoft yahei ui": ("msyhui", "msyh"),
    "simhei": ("simhei",),
    "simsun": ("simsun",),
    "songti sc": ("songti",),
    "times new roman": ("times", "timesnewroman"),
    "courier new": ("cour", "couriernew"),
}


# --------------------------------------------------------------------------
# font metrics
# --------------------------------------------------------------------------


class FontBook:
    """Advance widths in em units, read from real font files when available."""

    def __init__(self, dirs):
        self._dirs = [Path(d) for d in dirs if Path(d).is_dir()]
        self._cache = {}
        self.missing = set()

    def _candidates(self, family):
        key = family.strip().lower()
        stems = list(FONT_ALIASES.get(key, ()))
        stems.append(key.replace(" ", ""))
        stems.append(key)
        for directory in self._dirs:
            for stem in stems:
                for suffix in (".ttf", ".ttc", ".otf", ".TTF", ".TTC", ".OTF"):
                    candidate = directory / f"{stem}{suffix}"
                    if candidate.is_file():
                        yield candidate

    def _load(self, family):
        if TTFont is None:
            return None
        for path in self._candidates(family):
            try:
                font = (
                    TTCollection(str(path)).fonts[0]
                    if path.suffix.lower() == ".ttc"
                    else TTFont(str(path), fontNumber=0)
                )
                upem = font["head"].unitsPerEm
                return font.getBestCmap(), font["hmtx"].metrics, font.getGlyphOrder(), upem
            except Exception:
                continue
        return None

    def widths(self, family):
        """Return a callable char -> width in em, or None if unavailable."""
        if family not in self._cache:
            loaded = self._load(family)
            if loaded is None:
                self.missing.add(family)
            self._cache[family] = loaded
        loaded = self._cache[family]
        if loaded is None:
            return None
        cmap, hmtx, _order, upem = loaded

        def width(char):
            glyph = cmap.get(ord(char))
            if glyph is None or glyph not in hmtx:
                return None
            return hmtx[glyph][0] / upem

        return width


def _is_wide(char):
    return unicodedata.east_asian_width(char) in ("W", "F")


def text_width(text, size_pt, family, book):
    """Estimated advance width of `text` in inches."""
    width = book.widths(family) if family else None
    total_em = 0.0
    for char in text:
        if char == "\n":
            continue
        if _is_wide(char):
            total_em += 1.0            # CJK glyphs are one em in every CJK font
            continue
        measured = width(char) if width else None
        total_em += measured if measured is not None else LATIN_EM_RATIO
    return total_em * size_pt / 72.0


# --------------------------------------------------------------------------
# element model
# --------------------------------------------------------------------------


@dataclass
class Element:
    z: int
    name: str
    named: bool
    rect: tuple
    shape: object
    background: bool = False
    parent: int | None = None
    children: list = field(default_factory=list)

    @property
    def area(self):
        return (self.rect[2] - self.rect[0]) * (self.rect[3] - self.rect[1])

    @property
    def has_text(self):
        return self.shape.has_text_frame and bool(self.shape.text_frame.text.strip())

    def label(self):
        if self.has_text:
            snippet = self.shape.text_frame.text.strip().replace("\n", " ")[:16]
            return f'{self.name} "{snippet}"'
        return self.name


def _rect(shape):
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    left, top = Emu(shape.left).inches, Emu(shape.top).inches
    return (left, top, left + Emu(shape.width).inches, top + Emu(shape.height).inches)


def _fill_alpha(shape):
    """Opacity percent of the shape's own fill, or None when it has no solid fill."""
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return None
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        return None
    alpha = solid.find(f".//{qn('a:alpha')}")
    return 100.0 if alpha is None else int(alpha.get("val")) / 1000.0


def _has_visible_box(shape):
    """Whether a shape contributes visible area rather than only text geometry."""
    if shape.shape_type in {
        MSO_SHAPE_TYPE.CHART,
        MSO_SHAPE_TYPE.GROUP,
        MSO_SHAPE_TYPE.MEDIA,
        MSO_SHAPE_TYPE.PICTURE,
        MSO_SHAPE_TYPE.TABLE,
    }:
        return True
    if shape.shape_type == MSO_SHAPE_TYPE.LINE:
        return False
    spPr = shape._element.find(qn("p:spPr"))
    if spPr is None:
        return False
    if _fill_alpha(shape) is not None:
        return True
    if any(spPr.find(qn(tag)) is not None for tag in ("a:gradFill", "a:pattFill", "a:blipFill")):
        return True
    line = spPr.find(qn("a:ln"))
    if line is None or line.find(qn("a:noFill")) is not None:
        return False
    return any(
        line.find(qn(tag)) is not None
        for tag in ("a:solidFill", "a:gradFill", "a:pattFill")
    )


def contains(outer, inner, tol=TOL):
    return (
        outer[0] - tol <= inner[0]
        and outer[1] - tol <= inner[1]
        and outer[2] + tol >= inner[2]
        and outer[3] + tol >= inner[3]
    )


def overlap_area(a, b):
    width = min(a[2], b[2]) - max(a[0], b[0])
    height = min(a[3], b[3]) - max(a[1], b[1])
    if width <= OVERLAP_MIN or height <= OVERLAP_MIN:
        return 0.0      # a sliver this thin reads as two boxes merely abutting
    return width * height


def build_elements(slide):
    elements = []
    for z, shape in enumerate(slide.shapes):
        rect = _rect(shape)
        if rect is None:
            continue
        raw = (shape.name or "").strip()
        named = bool(raw) and not AUTO_NAME_RE.match(raw)
        name = raw if named else (raw or f"Shape {z}")
        background = named and (name == BG_PREFIX or name.startswith(BG_PREFIX + "/"))
        if not named:
            alpha = _fill_alpha(shape)
            if alpha is not None and alpha < INFERRED_BG_ALPHA:
                background = True   # unnamed translucent art is almost always a backdrop
        elements.append(Element(z=z, name=name, named=named, rect=rect,
                                shape=shape, background=background))
    _link(elements)
    return elements


def _link(elements):
    """Set parent/children. Declared paths win; unnamed elements are adopted."""
    by_name = {el.name: el for el in elements if el.named}

    for el in elements:
        if el.named and "/" in el.name:
            parent_name = el.name.rsplit("/", 1)[0]
            while parent_name:
                if parent_name in by_name:
                    el.parent = by_name[parent_name].z
                    break
                parent_name = parent_name.rsplit("/", 1)[0] if "/" in parent_name else ""
            if el.parent is not None:
                continue
        if el.named:
            continue

        # Unnamed: adopt into the smallest *named* container that holds it.
        best = None
        for other in elements:
            if other.z == el.z or not other.named or not contains(other.rect, el.rect):
                continue
            if best is None or other.area < best.area:
                best = other
        if best is not None and not contains(el.rect, best.rect):
            el.parent = best.z
        elif best is not None:
            # identical rectangles: the shape drawn later sits on the earlier one
            el.parent = best.z if best.z < el.z else None

    # Fully geometric fallback for decks with no declared names at all.
    if not by_name:
        for el in elements:
            best = None
            for other in elements:
                if other.z == el.z or not contains(other.rect, el.rect):
                    continue
                if contains(el.rect, other.rect) and other.z > el.z:
                    continue    # identical rects: later-drawn is the child
                if best is None or other.area < best.area:
                    best = other
            el.parent = best.z if best is not None else None

    index = {el.z: el for el in elements}
    for el in elements:
        if el.parent is not None and el.parent in index:
            index[el.parent].children.append(el.z)


def ancestors(elements, z):
    index = {el.z: el for el in elements}
    seen = set()
    current = index[z].parent
    while current is not None and current not in seen:
        seen.add(current)
        yield current
        current = index[current].parent if current in index else None


# --------------------------------------------------------------------------
# text geometry
# --------------------------------------------------------------------------


def _run_size(run, paragraph, shape):
    for source in (run.font.size, paragraph.font.size):
        if source is not None:
            return source.pt
    return DEFAULT_FONT_SIZE


def _run_family(run, paragraph):
    for source in (run.font.name, paragraph.font.name):
        if source:
            return source
    return None


def paragraph_metrics(shape, book):
    """(width_in, max_size_pt) per paragraph, in order."""
    out = []
    for paragraph in shape.text_frame.paragraphs:
        width = 0.0
        biggest = 0.0
        for run in paragraph.runs:
            size = _run_size(run, paragraph, shape)
            biggest = max(biggest, size)
            width += text_width(run.text, size, _run_family(run, paragraph), book)
        if paragraph.runs:
            out.append((width, biggest or DEFAULT_FONT_SIZE))
    return out


def inner_width(shape):
    frame = shape.text_frame
    left = Emu(frame.margin_left).inches if frame.margin_left is not None else 0.1
    right = Emu(frame.margin_right).inches if frame.margin_right is not None else 0.1
    return max(0.01, Emu(shape.width).inches - left - right)


def text_centre(el):
    """Vertical centre of the first text line, in inches."""
    frame = el.shape.text_frame
    top, bottom = el.rect[1], el.rect[3]
    sizes = [size for _w, size in paragraph_metrics(el.shape, _NULL_BOOK)] or [DEFAULT_FONT_SIZE]
    line = sizes[0] * LINE_HEIGHT / 72.0
    anchor = frame.vertical_anchor
    if anchor is not None and "MIDDLE" in str(anchor):
        return (top + bottom) / 2
    if anchor is not None and "BOTTOM" in str(anchor):
        inset = Emu(frame.margin_bottom).inches if frame.margin_bottom is not None else 0.05
        return bottom - inset - line / 2
    inset = Emu(frame.margin_top).inches if frame.margin_top is not None else 0.05
    return top + inset + line / 2


def _is_centred(el):
    frame = el.shape.text_frame
    anchor = frame.vertical_anchor
    if anchor is None or "MIDDLE" not in str(anchor):
        return False
    for paragraph in frame.paragraphs:
        if paragraph.alignment is not None and "CENTER" in str(paragraph.alignment):
            return True
    return False


class _NullBook:
    def widths(self, family):
        return None


_NULL_BOOK = _NullBook()


# --------------------------------------------------------------------------
# content balance
# --------------------------------------------------------------------------


def tight_text_rect(el, book):
    """Approximate the rendered text block instead of using its loose frame."""
    metrics = paragraph_metrics(el.shape, book)
    if not metrics:
        return None
    frame = el.shape.text_frame
    margin_left = Emu(frame.margin_left).inches if frame.margin_left is not None else 0.1
    margin_right = Emu(frame.margin_right).inches if frame.margin_right is not None else 0.1
    margin_top = Emu(frame.margin_top).inches if frame.margin_top is not None else 0.05
    margin_bottom = Emu(frame.margin_bottom).inches if frame.margin_bottom is not None else 0.05
    available_width = max(0.01, el.rect[2] - el.rect[0] - margin_left - margin_right)
    available_height = max(0.01, el.rect[3] - el.rect[1] - margin_top - margin_bottom)

    line_widths = []
    total_height = 0.0
    for width, size in metrics:
        line_count = max(1, math.ceil(width / available_width))
        remaining = width
        for _ in range(line_count):
            line_widths.append(min(available_width, remaining))
            remaining = max(0.0, remaining - available_width)
        total_height += line_count * size * LINE_HEIGHT / 72.0
    block_width = min(available_width, max(line_widths, default=available_width))
    block_height = min(available_height, total_height)

    alignments = [str(p.alignment) for p in frame.paragraphs if p.alignment is not None]
    if alignments and all("CENTER" in alignment for alignment in alignments):
        x1 = (el.rect[0] + el.rect[2] - block_width) / 2
    elif alignments and all("RIGHT" in alignment for alignment in alignments):
        x1 = el.rect[2] - margin_right - block_width
    else:
        x1 = el.rect[0] + margin_left

    anchor = str(frame.vertical_anchor) if frame.vertical_anchor is not None else ""
    if "MIDDLE" in anchor:
        y1 = (el.rect[1] + el.rect[3] - block_height) / 2
    elif "BOTTOM" in anchor:
        y1 = el.rect[3] - margin_bottom - block_height
    else:
        y1 = el.rect[1] + margin_top
    return (x1, y1, x1 + block_width, y1 + block_height)


def content_balance_metrics(rectangles, page_width, page_height):
    """Measure spatial occupancy without using slide or background colors."""
    left, top, right, bottom = CONTENT_BODY_BOUNDS
    body_x1, body_y1 = left * page_width, top * page_height
    body_x2, body_y2 = right * page_width, bottom * page_height
    body_width, body_height = body_x2 - body_x1, body_y2 - body_y1
    columns, rows = CONTENT_GRID_COLUMNS, CONTENT_GRID_ROWS
    occupied = [bytearray(columns) for _ in range(rows)]
    padding = 2.0 / 72.0

    for x1, y1, x2, y2 in rectangles:
        x1, y1 = max(body_x1, x1 - padding), max(body_y1, y1 - padding)
        x2, y2 = min(body_x2, x2 + padding), min(body_y2, y2 + padding)
        if x2 <= x1 or y2 <= y1:
            continue
        grid_x1 = max(0, int((x1 - body_x1) / body_width * columns))
        grid_x2 = min(columns, int((x2 - body_x1) / body_width * columns) + 1)
        grid_y1 = max(0, int((y1 - body_y1) / body_height * rows))
        grid_y2 = min(rows, int((y2 - body_y1) / body_height * rows) + 1)
        fill = b"\x01" * (grid_x2 - grid_x1)
        for row in range(grid_y1, grid_y2):
            occupied[row][grid_x1:grid_x2] = fill

    half_column = columns // 2
    half_row = rows // 2
    left_cells = sum(sum(row[:half_column]) for row in occupied)
    right_cells = sum(sum(row[half_column:]) for row in occupied)
    top_cells = sum(sum(row) for row in occupied[:half_row])
    bottom_cells = sum(sum(row) for row in occupied[half_row:])

    def longest_empty(values):
        longest = current = 0
        for value in values:
            current = 0 if value else current + 1
            longest = max(longest, current)
        return longest

    used_columns = [any(occupied[row][column] for row in range(rows)) for column in range(columns)]
    used_rows = [any(occupied[row]) for row in range(rows)]
    return {
        "left": 100.0 * left_cells / (rows * half_column),
        "right": 100.0 * right_cells / (rows * half_column),
        "top": 100.0 * top_cells / (half_row * columns),
        "bottom": 100.0 * bottom_cells / (half_row * columns),
        "empty_vertical": longest_empty(used_columns) / columns,
        "empty_horizontal": longest_empty(used_rows) / rows,
    }


def check_balance(slide_no, elements, page_width, page_height, book):
    text_chars = sum(len(el.shape.text_frame.text.strip()) for el in elements if el.has_text)
    if text_chars < CONTENT_MIN_TEXT_CHARS:
        return []
    slide_area = page_width * page_height
    rectangles = []
    footer_names = ("hd/page", "footer/", "foot/")
    for el in elements:
        name = el.name.lower()
        if el.background or name.startswith(footer_names) or el.area >= slide_area * 0.80:
            continue
        if _has_visible_box(el.shape):
            rectangles.append(el.rect)
        elif el.has_text:
            text_rect = tight_text_rect(el, book)
            if text_rect is not None:
                rectangles.append(text_rect)

    metrics = content_balance_metrics(rectangles, page_width, page_height)
    findings = []
    left, right = metrics["left"], metrics["right"]
    if (
        min(left, right) < CONTENT_EMPTY_HALF_PERCENT
        and max(left, right) > CONTENT_OCCUPIED_HALF_PERCENT
        and metrics["empty_vertical"] >= CONTENT_EMPTY_BAND_RATIO
    ):
        empty_side = "left" if left < right else "right"
        occupied_side = "right" if empty_side == "left" else "left"
        findings.append(Finding(
            slide=slide_no, kind="balance", waivable=False, severe=True,
            message=(f"{empty_side} body {min(left, right):.1f}% vs {occupied_side} "
                     f"{max(left, right):.1f}%; empty vertical band "
                     f"{metrics['empty_vertical']:.0%}"),
        ))

    top, bottom = metrics["top"], metrics["bottom"]
    if (
        min(top, bottom) < CONTENT_EMPTY_HALF_PERCENT
        and max(top, bottom) > CONTENT_OCCUPIED_HALF_PERCENT
        and metrics["empty_horizontal"] >= CONTENT_EMPTY_BAND_RATIO
    ):
        empty_side = "top" if top < bottom else "bottom"
        occupied_side = "bottom" if empty_side == "top" else "top"
        findings.append(Finding(
            slide=slide_no, kind="balance", waivable=False, severe=True,
            message=(f"{empty_side} body {min(top, bottom):.1f}% vs {occupied_side} "
                     f"{max(top, bottom):.1f}%; empty horizontal band "
                     f"{metrics['empty_horizontal']:.0%}"),
        ))
    return findings


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------


@dataclass
class Finding:
    slide: int
    kind: str
    message: str
    ratio: float = 0.0
    pair: tuple = ()
    waivable: bool = True
    severe: bool = False


def _root(elements, z):
    """Topmost ancestor of `z`, used to collapse one visual object into one finding."""
    chain = list(ancestors(elements, z))
    return chain[-1] if chain else z


def check_overlap(slide_no, elements):
    best = {}
    for i, low in enumerate(elements):
        if low.background:
            continue
        for high in elements[i + 1:]:
            if high.background:
                continue
            if low.z in set(ancestors(elements, high.z)) or high.z in set(ancestors(elements, low.z)):
                continue
            if contains(high.rect, low.rect) or contains(low.rect, high.rect):
                continue
            area = overlap_area(low.rect, high.rect)
            if not area:
                continue
            ratio = area / min(low.area, high.area)
            # A card and the text drawn on it are one object; report them once.
            key = tuple(sorted((_root(elements, low.z), _root(elements, high.z))))
            severe = ratio >= SEVERE_RATIO
            finding = Finding(
                slide=slide_no,
                kind="overlap",
                message=f"{low.label()} × {high.label()} — {ratio:.0%} covered",
                ratio=ratio,
                pair=(low.name, high.name),
                waivable=low.named and high.named and not severe,
                severe=severe,
            )
            if key not in best or ratio > best[key].ratio:
                best[key] = finding
    return list(best.values())


def check_escape(slide_no, elements):
    index = {el.z: el for el in elements}
    findings = []
    for el in elements:
        if el.parent is None or not el.named:
            continue
        parent = index.get(el.parent)
        if parent is None or not parent.named or parent.background:
            continue
        if contains(parent.rect, el.rect):
            continue
        out = max(
            parent.rect[0] - el.rect[0], el.rect[2] - parent.rect[2],
            parent.rect[1] - el.rect[1], el.rect[3] - parent.rect[3],
        )
        findings.append(Finding(
            slide=slide_no, kind="escape", waivable=False,
            message=f"{el.label()} sticks {out:.2f}in outside {parent.name}",
        ))
    return findings


def check_fit(slide_no, elements, book):
    findings = []
    for el in elements:
        if not el.has_text:
            continue
        frame = el.shape.text_frame
        if frame.auto_size is not None and "TEXT_TO_FIT" in str(frame.auto_size):
            continue          # shrink-to-fit boxes cannot overflow
        if not frame.word_wrap and frame.word_wrap is not None:
            continue
        available = inner_width(el.shape)
        metrics = paragraph_metrics(el.shape, book)
        if not metrics:
            continue
        needed = sum(max(1, math.ceil(width / available)) for width, _ in metrics)
        tallest = max(size for _w, size in metrics) * LINE_HEIGHT / 72.0
        capacity = max(1, math.floor((el.rect[3] - el.rect[1]) / tallest))
        if needed <= capacity:
            continue
        widest = max(width for width, _ in metrics)
        findings.append(Finding(
            slide=slide_no, kind="fit", waivable=False,
            message=(f"{el.label()} needs {needed} lines, box holds {capacity} "
                     f"(text {widest:.2f}in, inner width {available:.2f}in)"),
        ))
    return findings


def check_center(slide_no, elements):
    index = {el.z: el for el in elements}
    findings = []
    for el in elements:
        if not el.has_text or el.parent is None or not _is_centred(el):
            continue
        parent = index.get(el.parent)
        if parent is None or parent.has_text:
            continue
        dx = abs((el.rect[0] + el.rect[2]) / 2 - (parent.rect[0] + parent.rect[2]) / 2)
        dy = abs((el.rect[1] + el.rect[3]) / 2 - (parent.rect[1] + parent.rect[3]) / 2)
        if max(dx, dy) <= CENTER_TOL:
            continue
        findings.append(Finding(
            slide=slide_no, kind="center", waivable=False,
            message=(f"{el.label()} is centred but sits {dx:.2f}in / {dy:.2f}in "
                     f"off the centre of {parent.name}"),
        ))
    return findings


def check_baseline(slide_no, elements):
    index = {el.z: el for el in elements}
    groups = {}
    for el in elements:
        if el.has_text and el.parent is not None:
            groups.setdefault(el.parent, []).append(el)
    findings = []
    for parent_z, group in groups.items():
        parent = index.get(parent_z)
        if parent is None or parent.background:
            continue
        for i, a in enumerate(group):
            for b in group[i + 1:]:
                span = min(a.rect[3], b.rect[3]) - max(a.rect[1], b.rect[1])
                if span <= 0:
                    continue
                shorter = min(a.rect[3] - a.rect[1], b.rect[3] - b.rect[1])
                if shorter <= 0 or span / shorter < 0.5:
                    continue
                delta = abs(text_centre(a) - text_centre(b))
                if delta <= BASELINE_TOL:
                    continue
                findings.append(Finding(
                    slide=slide_no, kind="baseline", waivable=False,
                    message=(f"{a.label()} and {b.label()} share a row but their text "
                             f"centres differ by {delta:.2f}in"),
                ))
    return findings


def _point_touches_rect(point, rect, tol=CONNECTOR_TOL):
    x, y = point
    if not (rect[0] - tol <= x <= rect[2] + tol and rect[1] - tol <= y <= rect[3] + tol):
        return False
    return min(
        abs(x - rect[0]), abs(x - rect[2]), abs(y - rect[1]), abs(y - rect[3])
    ) <= tol


def check_connectors(slide_no, elements):
    """Require named flow arrows to touch a node at both endpoints."""
    directions = {
        "RIGHT_ARROW": lambda rect: (
            (rect[0], (rect[1] + rect[3]) / 2),
            (rect[2], (rect[1] + rect[3]) / 2),
        ),
        "LEFT_ARROW": lambda rect: (
            (rect[2], (rect[1] + rect[3]) / 2),
            (rect[0], (rect[1] + rect[3]) / 2),
        ),
        "DOWN_ARROW": lambda rect: (
            ((rect[0] + rect[2]) / 2, rect[1]),
            ((rect[0] + rect[2]) / 2, rect[3]),
        ),
        "UP_ARROW": lambda rect: (
            ((rect[0] + rect[2]) / 2, rect[3]),
            ((rect[0] + rect[2]) / 2, rect[1]),
        ),
    }
    findings = []
    for arrow in elements:
        if arrow.background or not arrow.named or not arrow.name.lower().startswith("flow/"):
            continue
        def _shape_type(sh):
            try:
                return str(sh.auto_shape_type).split()[0]
            except (ValueError, AttributeError):
                return ""
        arrow_type = _shape_type(arrow.shape)
        endpoints = directions.get(arrow_type)
        if endpoints is None:
            continue
        source, target = endpoints(arrow.rect)
        other_nodes = [
            node for node in elements
            if node.z != arrow.z and not node.background
            and not _shape_type(node.shape).endswith("_ARROW")
        ]
        source_attached = any(_point_touches_rect(source, node.rect) for node in other_nodes)
        target_attached = any(_point_touches_rect(target, node.rect) for node in other_nodes)
        missing = []
        if not source_attached:
            missing.append("source")
        if not target_attached:
            missing.append("target")
        if missing:
            findings.append(Finding(
                slide=slide_no, kind="connector", waivable=False, severe=True,
                message=(f"{arrow.name} {arrow_type} has detached "
                         f"{' and '.join(missing)} endpoint"),
            ))
    return findings


# --------------------------------------------------------------------------
# waivers
# --------------------------------------------------------------------------


def load_waivers(path):
    if not path or not Path(path).is_file():
        return []
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"WARNING  {path}: {exc}", file=sys.stderr)
        return []
    return data.get("waivers", []) if isinstance(data, dict) else []


def waiver_for(finding, waivers):
    """Return (waiver, reason_it_does_not_apply)."""
    for waiver in waivers:
        if waiver.get("slide") != finding.slide:
            continue
        if {waiver.get("a"), waiver.get("b")} != set(finding.pair):
            continue
        if not str(waiver.get("reason", "")).strip():
            return waiver, "waiver has no reason"
        if finding.severe:
            return waiver, f"overlap is at or above {SEVERE_RATIO:.0%} and cannot be waived"
        if not finding.waivable:
            return waiver, "waivers require both elements to be named"
        recorded = float(waiver.get("ratio", 0.0))
        if finding.ratio > recorded + WAIVER_SLACK:
            return waiver, (f"waiver expired: confirmed at {recorded:.0%}, "
                            f"now {finding.ratio:.0%}")
        return waiver, None
    return None, None


# --------------------------------------------------------------------------
# report
# --------------------------------------------------------------------------


def analyse(pptx_path, waivers_path=None, font_dirs=DEFAULT_FONT_DIRS):
    book = FontBook(font_dirs)
    waivers = load_waivers(waivers_path)
    lines = [f"LAYOUT  {pptx_path}"]
    total = severe_count = waived_count = 0
    named_total = element_total = 0
    presentation = Presentation(pptx_path)
    page_width = Emu(presentation.slide_width).inches
    page_height = Emu(presentation.slide_height).inches

    for slide_no, slide in enumerate(presentation.slides, 1):
        elements = build_elements(slide)
        named = sum(1 for el in elements if el.named)
        named_total += named
        element_total += len(elements)

        findings = (
            check_overlap(slide_no, elements)
            + check_escape(slide_no, elements)
            + check_fit(slide_no, elements, book)
            + check_center(slide_no, elements)
            + check_baseline(slide_no, elements)
            + check_connectors(slide_no, elements)
            + check_balance(slide_no, elements, page_width, page_height, book)
        )
        findings.sort(key=lambda f: (-f.ratio, f.kind))

        shown = []
        for finding in findings:
            waiver, blocked = waiver_for(finding, waivers)
            if waiver is not None and blocked is None:
                waived_count += 1
                continue
            shown.append((finding, blocked))

        lines.append(f"slide {slide_no:2} | {len(elements):3} elements | "
                     f"named {named} | inferred {len(elements) - named}")
        for finding, blocked in shown:
            total += 1
            tag = "  [not waivable]" if finding.severe else ""
            lines.append(f"  WARN  {finding.kind:<9} {finding.message}{tag}")
            if blocked:
                lines.append(f"        {blocked}")
            if finding.severe:
                severe_count += 1
            elif finding.kind == "overlap" and finding.waivable:
                waive = json.dumps({"slide": finding.slide, "a": finding.pair[0],
                                    "b": finding.pair[1],
                                    "ratio": round(finding.ratio, 2),
                                    "reason": "..."}, ensure_ascii=False)
                lines.append(f"        to waive: {waive}")

    if book.missing:
        lines.insert(1, "font metrics unavailable for "
                        + ", ".join(sorted(book.missing))
                        + " — width estimates are conservative")
    coverage = (named_total / element_total * 100) if element_total else 100.0
    lines.append("")
    lines.append(f"SUMMARY  {total} warnings ({severe_count} severe, {waived_count} waived)"
                 f" | naming coverage {coverage:.0f}% ({named_total}/{element_total})")
    return "\n".join(lines) + "\n"


def analyse_balance(pptx_path, font_dirs=DEFAULT_FONT_DIRS):
    book = FontBook(font_dirs)
    presentation = Presentation(pptx_path)
    page_width = Emu(presentation.slide_width).inches
    page_height = Emu(presentation.slide_height).inches
    lines = [f"CONTENT BALANCE  {pptx_path}"]
    total = 0
    for slide_no, slide in enumerate(presentation.slides, 1):
        findings = check_balance(
            slide_no, build_elements(slide), page_width, page_height, book
        )
        for finding in findings:
            total += 1
            lines.append(f"FAIL  slide {slide_no}: {finding.message}")
    if not total:
        lines.append(f"PASS  slides={len(presentation.slides)}; no severe content imbalance")
    return total == 0, "\n".join(lines) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("pptx")
    parser.add_argument("--waivers", default=None,
                        help="path to deck.layout.json")
    parser.add_argument("--font-dir", action="append", default=None,
                        help="extra directory to search for font files")
    parser.add_argument("--balance-only", action="store_true",
                        help="check only severe content imbalance and fail when found")
    args = parser.parse_args(argv)

    if not Path(args.pptx).is_file():
        parser.error(f"no such file: {args.pptx}")
    dirs = tuple(args.font_dir or ()) + DEFAULT_FONT_DIRS
    if args.balance_only:
        passed, report = analyse_balance(args.pptx, dirs)
        sys.stdout.write(report)
        return 0 if passed else 1
    sys.stdout.write(analyse(args.pptx, args.waivers, dirs))
    return 0        # advisory only: never fail the caller's build


if __name__ == "__main__":
    sys.exit(main())
