"""Find layout defects in a .pptx from shape geometry and font metrics.

Reports occlusion, crowding, text overflow or underfill, alignment defects and
severe content imbalance without rendering the deck, so it can be run while the
deck is still being built. It is the substitute for visual QA when image
inspection is unavailable.

Usage:
    python layout_check.py deck.pptx [--font-dir DIR]

Examples:
    python layout_check.py deck.pptx
    python layout_check.py deck.pptx --font-dir ~/fonts

The rule
    Any two elements must either nest completely, or stand CLEARANCE apart.
    Anything else — a partial overlap, or a hairline gap — is a defect.

    Elements are judged by the cells they ink, not by their bounding boxes: an
    ellipse is an ellipse, and a label is its lines of text rather than the
    loose frame around them. A text frame may therefore hang far outside its
    card without complaint, so long as the glyphs stay inside.

Backing
    Geometry alone cannot tell an element that landed outside the card it was
    meant to sit on from one that was always meant to sit on the slide: nothing
    holds it either way, so there is nothing for it to have escaped from. Only
    the generator knows, so it says so, in the name:

        objectName: "points"        a card, sitting on the slide
        objectName: "points/num4"   sits on `points`, and must stay on it
        objectName: "@bg/blob"      background art, exempt

    The text left of the last `/` must be the full objectName of another element
    on the same slide. An element whose name has no `/` is declaring that it
    sits on the slide itself. A declaration that cannot be resolved — no such
    element, or two elements answering to the name — is reported rather than
    skipped: a check that could not be evaluated has not passed.
"""

import argparse
import collections
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
CENTER_TOL = 0.01       # inches; concentricity slack for centred text
BASELINE_TOL = 0.05     # inches; sibling text centreline slack
SEVERE_MOVE = 0.25      # inches; a clearing move this big means something is buried
GRID_DPI = 100          # mask cells per inch; 0.05in clearance is exactly 5 cells
CLEARANCE = 0.05        # inches; unrelated elements must stand this far apart
STROKE_EPS = 0.001      # inches; thinner than this in either axis is a stroke
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
TEXT_UNDERFILL_MIN_HEIGHT = 0.60
TEXT_UNDERFILL_MIN_CHARS = 20
TEXT_UNDERFILL_MIN_RATIO = 0.40

CHECK_HINTS = {
    "overlap": ("Move or resize the elements until they no longer overlap. If one "
                "is intended as a container, make it fully contain the other."),
    "clearance": (f"Increase the gap to at least {CLEARANCE:.2f}in. Move aligned "
                  "peer elements together when needed to preserve their rhythm."),
    "fit": ("Enlarge the text box, shorten the text, or reduce the font size. If "
            "changing font size, apply the same change to peer boxes in the same group."),
    "leading": ("Increase line spacing to at least the font size. Keep line spacing "
                "and font size consistent across peer text boxes."),
    "center": ("Align the text box and its parent to the same centre, or remove centred "
               "alignment when the offset is intentional."),
    "baseline": ("Give text boxes in the row consistent y positions, heights, and "
                 "vertical anchors."),
    "balance": ("Redistribute or resize existing content, or add relevant content to "
                "the sparse region; do not fill the gap with decorative clutter."),
    "underfill": ("Increase the font size or add more text content. If changing font "
                  "size, apply the same change to peer boxes in the same group."),
    "backing": ("Move or resize the element until it sits entirely on the backing it "
                "declares, or enlarge that backing. If it was never meant to sit there, "
                "rename it after the backing it does sit on."),
    "dangling": ("Name the backing element exactly as the prefix spells it, or drop the "
                 "prefix when the element sits directly on the slide. `parent/child` "
                 "declares a backing; it is not a namespace."),
    "ambiguous": ("Give the elements distinct objectNames. A backing is resolved by name "
                  "alone — resolving it by geometry instead would make the check unable "
                  "to fail."),
}

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
    def is_stroke(self):
        """A line or rule: it has no interior, so nothing can be inside it.

        Chained polyline segments have to meet, and a segment has to run into
        its own arrowhead, so strokes take no part in occlusion or clearance.
        """
        return (self.rect[2] - self.rect[0] < STROKE_EPS
                or self.rect[3] - self.rect[1] < STROKE_EPS)

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


def _styled(shape, tag):
    """Whether `p:style` supplies a fill or a line. idx="0" means none."""
    style = shape._element.find(qn("p:style"))
    ref = style.find(qn(tag)) if style is not None else None
    return ref is not None and ref.get("idx") not in (None, "0")


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
    if spPr.find(qn("a:noFill")) is None:
        if _fill_alpha(shape) is not None:
            return True
        if any(spPr.find(qn(tag)) is not None
               for tag in ("a:gradFill", "a:pattFill", "a:blipFill")):
            return True
        if _styled(shape, "a:fillRef"):
            return True     # a themed shape keeps its fill in p:style, not spPr
    line = spPr.find(qn("a:ln"))
    if line is None:
        return _styled(shape, "a:lnRef")
    if line.find(qn("a:noFill")) is not None:
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


def separation(a, b):
    """Chebyshev gap between two rects in inches; <= 0 when they overlap."""
    return max(max(a[0] - b[2], b[0] - a[2]), max(a[1] - b[3], b[1] - a[3]))


def clearing_move(a, b):
    """Shortest move that would pull two overlapping rects apart."""
    return min(a[2] - b[0], b[2] - a[0], a[3] - b[1], b[3] - a[1])


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
    """Parent every element to the smallest element that geometrically holds it.

    Declared `parent/child` paths used to decide this, and used to exempt the
    pair from the occlusion check. Both jobs are now done by the geometry
    itself, so a name carries no structure — only a readable label.
    """
    for el in elements:
        best = None
        for other in elements:
            if other.z == el.z or not contains(other.rect, el.rect):
                continue
            if contains(el.rect, other.rect) and other.z > el.z:
                continue    # identical rects: the later-drawn one is the child
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
# geometry masks
# --------------------------------------------------------------------------


class Canvas:
    """A fixed cell grid for one slide.

    Every mask is a plain integer over the same grid, so the questions this
    module asks of two shapes — do they intersect, does one hold the other, how
    far apart are they — are single bit operations, and they are asked of the
    real outline rather than of a bounding box.
    """

    def __init__(self, width, height, dpi=GRID_DPI):
        self.dpi = dpi
        self.cols = max(1, math.ceil(width * dpi))
        self.rows = max(1, math.ceil(height * dpi))
        self.full = (1 << (self.rows * self.cols)) - 1
        first = sum(1 << (row * self.cols) for row in range(self.rows))
        self.not_first = self.full ^ first
        self.not_last = self.full ^ (first << (self.cols - 1))

    def _col(self, x):
        return min(self.cols, max(0, math.floor(x * self.dpi + 0.5)))

    def _row(self, y):
        return min(self.rows, max(0, math.floor(y * self.dpi + 0.5)))

    def _spans(self, spans):
        mask = 0
        for row, c1, c2 in spans:
            if 0 <= row < self.rows and c2 > c1:
                mask |= ((1 << (c2 - c1)) - 1) << (row * self.cols + c1)
        return mask

    def rect(self, rect):
        x1, y1, x2, y2 = rect
        c1, c2 = self._col(x1), self._col(x2)
        return self._spans(
            (row, c1, c2) for row in range(self._row(y1), self._row(y2))
        )

    def ellipse(self, rect):
        x1, y1, x2, y2 = rect
        centre_x, centre_y = (x1 + x2) / 2, (y1 + y2) / 2
        radius_x, radius_y = (x2 - x1) / 2, (y2 - y1) / 2
        if radius_x <= 0 or radius_y <= 0:
            return 0
        spans = []
        for row in range(self._row(y1), self._row(y2)):
            offset = ((row + 0.5) / self.dpi - centre_y) / radius_y
            if abs(offset) >= 1:
                continue
            half = radius_x * math.sqrt(1 - offset * offset)
            spans.append((row, self._col(centre_x - half), self._col(centre_x + half)))
        return self._spans(spans)

    def polygon(self, points):
        tops = [point[1] for point in points]
        spans = []
        for row in range(self._row(min(tops)), self._row(max(tops))):
            y = (row + 0.5) / self.dpi
            crossings = []
            for (ax, ay), (bx, by) in zip(points, points[1:] + points[:1]):
                if (ay <= y < by) or (by <= y < ay):
                    crossings.append(ax + (y - ay) * (bx - ax) / (by - ay))
            crossings.sort()
            for i in range(0, len(crossings) - 1, 2):
                spans.append((row, self._col(crossings[i]), self._col(crossings[i + 1])))
        return self._spans(spans)

    def grow(self, mask, inches):
        """Grow a mask outward by `inches`, Chebyshev — a square kernel."""
        for _ in range(max(0, int(inches * self.dpi + 0.5))):
            mask |= ((mask & self.not_last) << 1) | ((mask & self.not_first) >> 1)
            mask |= (mask << self.cols) | (mask >> self.cols)
            mask &= self.full
        return mask


def mask_holds(outer, inner):
    """Whether `outer` covers every cell of `inner`."""
    return inner | outer == outer      # `outer & ~inner` would go negative


def _preset(shape):
    spPr = shape._element.find(qn("p:spPr"))
    geometry = spPr.find(qn("a:prstGeom")) if spPr is not None else None
    return geometry.get("prst") if geometry is not None else None


def element_mask(el, canvas, book):
    """The cells `el` actually inks. Strokes and empty frames ink nothing."""
    if el.is_stroke:
        return 0
    if _has_visible_box(el.shape):
        preset = _preset(el.shape)
        if preset == "ellipse":
            return canvas.ellipse(el.rect)
        x1, y1, x2, y2 = el.rect
        if preset in ("triangle", "isocelesTriangle"):
            return canvas.polygon([((x1 + x2) / 2, y1), (x2, y2), (x1, y2)])
        return canvas.rect(el.rect)
    if el.has_text:
        mask = 0
        for line in text_lines(el, book):
            mask |= canvas.rect(line)
        return mask
    return 0


def element_ink(el, book):
    """Bounding rectangle of what `el` inks, for measuring distances in inches."""
    if el.is_stroke:
        return None
    if _has_visible_box(el.shape):
        return el.rect
    if el.has_text:
        lines = text_lines(el, book)
        if lines:
            return (min(line[0] for line in lines), lines[0][1],
                    max(line[2] for line in lines), lines[-1][3])
    return None


# --------------------------------------------------------------------------
# text geometry
# --------------------------------------------------------------------------


def _run_size(run, paragraph, shape):
    for source in (run.font.size, paragraph.font.size):
        if source is not None:
            return source.pt
    return DEFAULT_FONT_SIZE


def line_height(paragraph, size_pt):
    """Rendered distance between successive baselines, in points.

    `a:lnSpc` overrides the default: `a:spcPts` states it outright in hundredths
    of a point, `a:spcPct` as a percentage of single spacing.
    """
    properties = paragraph._p.find(qn("a:pPr"))
    spacing = properties.find(qn("a:lnSpc")) if properties is not None else None
    if spacing is not None:
        points = spacing.find(qn("a:spcPts"))
        if points is not None:
            return int(points.get("val")) / 100.0
        percent = spacing.find(qn("a:spcPct"))
        if percent is not None:
            return size_pt * LINE_HEIGHT * int(percent.get("val")) / 100000.0
    return size_pt * LINE_HEIGHT


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


def text_lines(el, book):
    """One rectangle per rendered line of text, in inches.

    Ragged text covers far less of its frame than the frame suggests, and the
    difference decides whether two labels collide. Line breaking is modelled as
    even packing (`ceil(width / available)`), so a line that would really break
    early on a word boundary is estimated a little narrow.
    """
    metrics = paragraph_metrics(el.shape, book)
    if not metrics:
        return []
    frame = el.shape.text_frame
    margin_left = Emu(frame.margin_left).inches if frame.margin_left is not None else 0.1
    margin_right = Emu(frame.margin_right).inches if frame.margin_right is not None else 0.1
    margin_top = Emu(frame.margin_top).inches if frame.margin_top is not None else 0.05
    margin_bottom = Emu(frame.margin_bottom).inches if frame.margin_bottom is not None else 0.05
    available_width = max(0.01, el.rect[2] - el.rect[0] - margin_left - margin_right)
    available_height = max(0.01, el.rect[3] - el.rect[1] - margin_top - margin_bottom)

    lines = []          # (width, height) per rendered line, in order
    for width, size in metrics:
        line_count = max(1, math.ceil(width / available_width))
        remaining = width
        for _ in range(line_count):
            lines.append((min(available_width, remaining), size * LINE_HEIGHT / 72.0))
            remaining = max(0.0, remaining - available_width)
    block_height = min(available_height, sum(height for _w, height in lines))

    alignments = [str(p.alignment) for p in frame.paragraphs if p.alignment is not None]
    centred = bool(alignments) and all("CENTER" in a for a in alignments)
    right = bool(alignments) and all("RIGHT" in a for a in alignments)

    anchor = str(frame.vertical_anchor) if frame.vertical_anchor is not None else ""
    if "MIDDLE" in anchor:
        y = (el.rect[1] + el.rect[3] - block_height) / 2
    elif "BOTTOM" in anchor:
        y = el.rect[3] - margin_bottom - block_height
    else:
        y = el.rect[1] + margin_top

    out = []
    for width, height in lines:
        if centred:
            x = (el.rect[0] + el.rect[2] - width) / 2
        elif right:
            x = el.rect[2] - margin_right - width
        else:
            x = el.rect[0] + margin_left
        out.append((x, y, x + width, y + height))
        y += height
    return out


def tight_text_rect(el, book):
    """Bounding box of the rendered text, instead of its loose frame."""
    lines = text_lines(el, book)
    if not lines:
        return None
    frame = el.shape.text_frame
    margin_top = Emu(frame.margin_top).inches if frame.margin_top is not None else 0.05
    margin_bottom = Emu(frame.margin_bottom).inches if frame.margin_bottom is not None else 0.05
    available_height = max(0.01, el.rect[3] - el.rect[1] - margin_top - margin_bottom)
    x1 = min(line[0] for line in lines)
    x2 = max(line[2] for line in lines)
    y1 = lines[0][1]
    return (x1, y1, x2, y1 + min(available_height, lines[-1][3] - y1))


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
            slide=slide_no, kind="balance", severe=True,
            message=(f"{empty_side} body {min(left, right):.1f}% vs {occupied_side} "
                     f"{max(left, right):.1f}%; empty vertical band "
                     f"{metrics['empty_vertical']:.0%} — clears when the empty half "
                     f"reaches {CONTENT_EMPTY_HALF_PERCENT:.0f}% or the band drops "
                     f"under {CONTENT_EMPTY_BAND_RATIO:.0%}"),
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
            slide=slide_no, kind="balance", severe=True,
            message=(f"{empty_side} body {min(top, bottom):.1f}% vs {occupied_side} "
                     f"{max(top, bottom):.1f}%; empty horizontal band "
                     f"{metrics['empty_horizontal']:.0%} — clears when the empty half "
                     f"reaches {CONTENT_EMPTY_HALF_PERCENT:.0f}% or the band drops "
                     f"under {CONTENT_EMPTY_BAND_RATIO:.0%}"),
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
    move: float = 0.0      # inches something must shift to satisfy the rule
    severe: bool = False


def _root(elements, z):
    """Topmost ancestor of `z`, used to collapse one visual object into one finding."""
    chain = list(ancestors(elements, z))
    return chain[-1] if chain else z


def check_overlap(slide_no, elements, canvas, book):
    """Every pair must either nest completely or stand `CLEARANCE` apart.

    Nothing here reads names. An element is judged by the cells it inks, and ink
    means the real outline — an ellipse is an ellipse, and a label is its lines
    of text rather than the loose frame drawn around them. A child that pokes
    out of its parent is therefore just an overlap; there is no separate escape
    check, and no way to declare a collision intentional.
    """
    live = []
    for el in elements:
        if el.background or el.is_stroke:
            continue
        mask = element_mask(el, canvas, book)
        if mask:
            live.append((el, mask, element_ink(el, book)))

    # One cell of slack absorbs the grid's rounding at a flush inner edge.
    padded = [canvas.grow(mask, 1 / canvas.dpi) for _el, mask, _ink in live]
    near = [canvas.grow(mask, CLEARANCE) for _el, mask, _ink in live]

    best = {}
    for i, (low, low_mask, low_ink) in enumerate(live):
        for j in range(i + 1, len(live)):
            high, high_mask, high_ink = live[j]
            if mask_holds(padded[i], high_mask) or mask_holds(padded[j], low_mask):
                continue                            # one nests inside the other
            if low_mask & high_mask:
                move = clearing_move(low_ink, high_ink)
                kind = "overlap"
                detail = f"overlapping, move {move:.2f}in to clear"
            elif near[i] & high_mask:
                gap = max(0.0, separation(low_ink, high_ink))
                move = CLEARANCE - gap
                kind = "clearance"
                detail = f"{gap:.2f}in apart, needs {CLEARANCE:.2f}in"
            else:
                continue
            # A card and the text drawn on it are one object; report them once.
            key = tuple(sorted((_root(elements, low.z), _root(elements, high.z))))
            finding = Finding(
                slide=slide_no,
                kind=kind,
                message=f"{low.label()} × {high.label()} — {detail}",
                move=move,
                severe=move >= SEVERE_MOVE,
            )
            if key not in best or move > best[key].move:
                best[key] = finding
    return list(best.values())


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
            slide=slide_no, kind="fit",
            message=(f"{el.label()} needs {needed} lines, box holds {capacity} "
                     f"(text {widest:.2f}in, inner width {available:.2f}in)"),
        ))
    return findings


def check_underfill(slide_no, elements, book):
    """Warn when body text occupies too little of a tall, top-anchored frame.

    The vertical ratio targets the visually hollow cards this check is meant to
    catch. A two-dimensional ink-area ratio would also flag ordinary ragged
    paragraphs, while centred banners and compact labels are intentional and
    therefore excluded.
    """
    findings = []
    for el in elements:
        if not el.has_text:
            continue
        frame = el.shape.text_frame
        anchor = str(frame.vertical_anchor) if frame.vertical_anchor is not None else ""
        if "MIDDLE" in anchor or "BOTTOM" in anchor:
            continue
        text = frame.text.strip()
        if len(text) < TEXT_UNDERFILL_MIN_CHARS:
            continue

        margin_top = Emu(frame.margin_top).inches if frame.margin_top is not None else 0.05
        margin_bottom = Emu(frame.margin_bottom).inches if frame.margin_bottom is not None else 0.05
        inner_height = el.rect[3] - el.rect[1] - margin_top - margin_bottom
        if inner_height < TEXT_UNDERFILL_MIN_HEIGHT:
            continue
        lines = text_lines(el, book)
        if not lines:
            continue
        text_height = lines[-1][3] - lines[0][1]
        ratio = text_height / inner_height
        if ratio >= TEXT_UNDERFILL_MIN_RATIO:
            continue
        findings.append(Finding(
            slide=slide_no, kind="underfill",
            message=(f"{el.label()} uses {ratio:.0%} of its {inner_height:.2f}in "
                     f"inner height; expected at least {TEXT_UNDERFILL_MIN_RATIO:.0%}"),
        ))
    return findings


def check_leading(slide_no, elements, book):
    """Lines must stand at least as far apart as the glyphs are tall.

    Line spacing under the font size stacks every line on the one above it,
    which no other check sees: the text still fits its box — better than ever,
    in fact, since the lines pile up — and its ink still clears its neighbours.
    """
    findings = []
    for el in elements:
        if not el.has_text:
            continue
        available = inner_width(el.shape)
        for paragraph in el.shape.text_frame.paragraphs:
            if not paragraph.runs:
                continue
            size = max(_run_size(run, paragraph, el.shape) for run in paragraph.runs)
            width = sum(
                text_width(run.text, _run_size(run, paragraph, el.shape),
                           _run_family(run, paragraph), book)
                for run in paragraph.runs
            )
            if math.ceil(width / available) < 2:
                continue        # one line cannot collide with itself
            leading = line_height(paragraph, size)
            if leading >= size:
                continue
            findings.append(Finding(
                slide=slide_no, kind="leading", severe=True,
                message=(f"{el.label()} — {leading:.2f}pt line spacing under "
                         f"{size:.1f}pt text, so every line lands on the one above"),
            ))
            break               # one report per element is enough
    return findings


def check_center(slide_no, elements):
    index = {el.z: el for el in elements}
    findings = []
    for el in elements:
        if not el.has_text or el.parent is None or not _is_centred(el):
            continue
        parent = index.get(el.parent)
        if parent is None or parent.has_text or parent.background:
            continue        # being off the centre of the slide backdrop is not a defect
        dx = abs((el.rect[0] + el.rect[2]) / 2 - (parent.rect[0] + parent.rect[2]) / 2)
        dy = abs((el.rect[1] + el.rect[3]) / 2 - (parent.rect[1] + parent.rect[3]) / 2)
        if max(dx, dy) <= CENTER_TOL:
            continue
        findings.append(Finding(
            slide=slide_no, kind="center",
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
                    slide=slide_no, kind="baseline",
                    message=(f"{a.label()} and {b.label()} share a row but their text "
                             f"centres differ by {delta:.2f}in, over the "
                             f"{BASELINE_TOL:.2f}in allowed"),
                ))
    return findings


def check_backing(slide_no, elements):
    """Every element sits entirely on the backing its name declares.

    This is the one thing geometry cannot work out for itself. An element that
    missed its card is held by nothing, so no containment test has anything to
    compare it against; its own name is the only record of where it belonged.

    The declared backing is resolved by name and by name only. Picking the
    candidate that happens to contain the element would satisfy the rule by
    construction and the check could never fail.
    """
    seen = collections.Counter(el.name for el in elements if el.named)
    by_name = {el.name: el for el in elements if el.named}
    findings = []
    for el in elements:
        if el.background or not el.named or "/" not in el.name:
            continue        # decoration, an auto name declares nothing, or on the slide
        claim = el.name.rsplit("/", 1)[0]
        if seen[claim] > 1:
            findings.append(Finding(
                slide=slide_no, kind="ambiguous",
                message=(f"{el.label()} declares backing {claim!r}, but "
                         f"{seen[claim]} elements on this slide answer to that name"),
            ))
            continue
        if not seen[claim]:
            findings.append(Finding(
                slide=slide_no, kind="dangling",
                message=(f"{el.label()} declares backing {claim!r}, which is not an "
                         f"element on this slide, so it was never checked"),
            ))
            continue
        backing = by_name[claim]
        if contains(backing.rect, el.rect):
            continue
        out = max(backing.rect[0] - el.rect[0], el.rect[2] - backing.rect[2],
                  backing.rect[1] - el.rect[1], el.rect[3] - backing.rect[3])
        findings.append(Finding(
            slide=slide_no, kind="backing", move=out, severe=out >= SEVERE_MOVE,
            message=f"{el.label()} sits {out:.2f}in outside its backing {claim}",
        ))
    return findings



# --------------------------------------------------------------------------
# report
# --------------------------------------------------------------------------


def analyse(pptx_path, font_dirs=DEFAULT_FONT_DIRS):
    book = FontBook(font_dirs)
    lines = [f"LAYOUT  {pptx_path}"]
    total = severe_count = 0
    named_total = element_total = 0
    presentation = Presentation(pptx_path)
    page_width = Emu(presentation.slide_width).inches
    page_height = Emu(presentation.slide_height).inches
    canvas = Canvas(page_width, page_height)

    for slide_no, slide in enumerate(presentation.slides, 1):
        elements = build_elements(slide)
        named = sum(1 for el in elements if el.named)
        named_total += named
        element_total += len(elements)

        findings = (
            check_overlap(slide_no, elements, canvas, book)
            + check_fit(slide_no, elements, book)
            + check_underfill(slide_no, elements, book)
            + check_leading(slide_no, elements, book)
            + check_center(slide_no, elements)
            + check_baseline(slide_no, elements)
            + check_backing(slide_no, elements)
            + check_balance(slide_no, elements, page_width, page_height, book)
        )
        findings.sort(key=lambda f: (-f.move, f.kind))

        lines.append(f"slide {slide_no:2} | {len(elements):3} elements | "
                     f"named {named} | inferred {len(elements) - named}")
        hinted = set()
        for finding in findings:
            total += 1
            tag = "  [severe]" if finding.severe else ""
            lines.append(f"  WARN  {finding.kind:<9} {finding.message}{tag}")
            if finding.kind not in hinted:
                hint = CHECK_HINTS.get(finding.kind)
                if hint:
                    lines.append(f"        HINT  {hint}")
                hinted.add(finding.kind)
            if finding.severe:
                severe_count += 1

    if book.missing:
        lines.insert(1, "font metrics unavailable for "
                        + ", ".join(sorted(book.missing))
                        + " — width estimates are conservative")
    coverage = (named_total / element_total * 100) if element_total else 100.0
    lines.append("")
    lines.append(f"SUMMARY  {total} warnings ({severe_count} severe)"
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
        if findings:
            lines.append(f"      HINT  {CHECK_HINTS['balance']}")
    if not total:
        lines.append(f"PASS  slides={len(presentation.slides)}; no severe content imbalance")
    return total == 0, "\n".join(lines) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("pptx")
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
    sys.stdout.write(analyse(args.pptx, dirs))
    return 0        # advisory only: never fail the caller's build


if __name__ == "__main__":
    sys.exit(main())
