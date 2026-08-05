# PPT Harness

A deliberately small Python agent loop that asks DeepSeek to build an editable
PowerPoint presentation. It loads the bundled official Anthropic PPTX skill
from `./pptx/SKILL.md`, exposes four local tools (`read_file`, `write_file`, `edit_file`,
and `bash`), and performs deterministic QA before accepting the result.

DeepSeek V4 is text-only. The harness renders slides and creates a contact sheet for
human review, but it does not send images to the model or pretend that the model
performed visual inspection. [Layout QA](#layout-qa) recovers collisions, overflow,
and misalignment from coordinates rather than pixels. Layout QA also checks severe
content imbalance without depending on background color: it combines tight text
bounds estimated from font metrics with visible PPT object bounds and rejects large
unused body regions.
Subtler density, contrast, and type-hierarchy judgments still rest on the skill's
visual-QA checklist.

The bundled skill is therefore the official one plus those local edits, so its
sha256 — logged to `run.jsonl` on every run — no longer matches any upstream
release. Diff against the upstream copy before resyncing: `git diff` on
`pptx/SKILL.md` shows exactly what is local.

## Requirements

- Python 3.11+
- Node.js 18+ (only when the agent chooses PptxGenJS)
- LibreOffice, Poppler (`pdfinfo`, `pdftotext`, `pdftoppm`), ImageMagick `montage`,
  and `unzip`

Install both supported generator stacks:

```bash
python3 -m pip install -r requirements.txt
npm install
```

The official skill is authoritative. For a new deck it directs the model to use
PptxGenJS and write `deck.js`; its template-editing rules may select lower-level
OOXML tools. Every path must produce `deck.pptx`.

## Usage

```bash
export DEEPSEEK_API_KEY='your-key'

python3 ppt_agent.py --out ./runs/demo \
  '制作一份面向非技术员工的 10 页大语言模型幻觉科普演示文稿'
```

The request can also come from stdin:

```bash
printf '%s\n' 'Create a five-slide project update' | \
  python3 ppt_agent.py --out ./runs/project-update
```

Options:

```text
--out PATH          Required run directory
--model NAME        Defaults to DEEPSEEK_MODEL or deepseek-v4-flash
--max-steps NUMBER  Defaults to 40
--skill PATH        Defaults to ./pptx/SKILL.md
```

`DEEPSEEK_BASE_URL` overrides the default
`https://api.deepseek.com/anthropic` endpoint.

## Output

Each run keeps its source and audit trail together:

```text
runs/demo/
├── deck.py or deck.js
├── deck.pptx
├── deck.pdf
├── preview.jpg
├── qa.txt
├── run.jsonl
└── slides/slide-*.jpg
```

When the agent stops calling tools, the harness runs the skill's official
`validate.py`, MarkItDown checks, sandbox-compatible `soffice.py`, page-count and
text-boundary checks, PDF rendering, and preview generation. QA uses a temporary
directory and only replaces `deck.pdf`, `slides/`, and `preview.jpg` after every
check passes. The contact sheet is for human visual review.

The `content_balance` check in `pptx/scripts/layout_check.py` divides the slide body
into a fine occupancy grid. Text uses tight bounds estimated from its font metrics
rather than the often oversized text frame; filled or outlined objects, images,
charts, tables, and groups use their OOXML bounds. Background objects, page numbers,
and invisible text boxes do not count. A page fails only for a severe case: one half
below 5% occupancy, the other above 25%, plus a continuous empty band of at least 25%
of the body. This catches one-sided checklist layouts without relying on a white or
solid-color background. Use `--balance-only` when this metric should be a blocking
verification instead of an advisory layout report.

## Layout QA

`pptx/scripts/layout_check.py` finds the defects a text-only model cannot see:
one element covering another, elements crowded against each other, text too wide
for its box, and elements that should share a centreline but don't. It reads
shape coordinates and font metrics — nothing is
rendered, so the model runs it while it builds rather than only at the end.

### The rule

Any two elements must either **nest completely** or **stand 0.05" apart**.
A partial overlap is a defect, and so is a hairline gap. There is no third
category and no way to declare an exception, so every warning means something
has to move.

Elements are judged by the cells they ink rather than by their bounding boxes.
Each slide is rasterised onto a 0.01" grid and every shape becomes a bitmask, so
an ellipse is an ellipse and a label is its lines of text, not the loose frame
around them — a text frame may hang well outside its card without complaint as
long as the glyphs stay inside. Distances are Chebyshev, so two elements offset
diagonally need only clear each other on one axis.

Severity is the distance something must move: a 0.03" overlap is a nudge, a 1"
overlap means an element is buried. Coverage ratio cannot tell those apart — it
calls both of them nearly 100%.

Lines and rules take no part in this. A stroke has no interior, and chained
polyline segments have to meet, so they are excluded from both halves of the
rule.

The rule is strict on purpose, but the output is advisory. It measures geometry and
never sees a rendering — a chart is one solid rectangle to it, a rotated element is
judged by its unrotated box, line breaking is estimated rather than laid out — so a
warning marks a place to look, not a proven defect. The skill asks the model to
account for every warning, by fixing it or by saying in its final report why it
didn't, and gives it no way to suppress one. That is why removing waivers cost
nothing: a waiver file is only necessary when the output is treated as a verdict.

```bash
python3 pptx/scripts/layout_check.py runs/demo/deck.pptx
```

It is advisory: it always exits 0 and never blocks QA. Nothing forces the model to
run it, so `grep layout_check runs/<name>/run.jsonl` is how you check that it did.

### Naming

Names no longer affect any verdict. `parent/child` paths once declared that an
overlap was intentional, and a waiver file could excuse the rest; both are gone,
because the geometry already says everything the paths were claiming. What a name
buys now is a readable report — `card1/title` beats `Shape 7` when you are
reading a warning — plus one special case:

```js
s.addText("标题",       { ..., objectName: "card1/title" });  // just a label
s.addShape("ellipse",  { ..., objectName: "@bg/blob" });     // background art
```

`@bg` marks decoration that anything may cover. That is the only prefix the check
reads.

The cost of dropping the escape hatches is real: a deliberate partial overlap — a
badge straddling a card's corner, a callout hanging over the bar it marks — can no
longer be expressed at all, and has to be redrawn as a nested or a separated pair.

### Fonts

Width estimates are exact for CJK text (one em per glyph) and need the real font
file for Latin text. Point the checker at one with `--font-dir`; `/mnt/c/Windows/Fonts`,
`/usr/share/fonts`, and `/Library/Fonts` are searched by default. When a font is
missing, the report says so and falls back to a conservative estimate.

Installing the deck's fonts also makes the rendered `slides/*.jpg` match real
PowerPoint. On WSL:

```bash
echo '<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>
  <dir>/mnt/c/Windows/Fonts</dir></fontconfig>' > ~/.config/fontconfig/fonts.conf
fc-cache -f
```

## Security

The `bash` tool intentionally has unrestricted local shell access. Run this harness
only on a trusted machine and with prompts you trust. API credential variables are
removed from the tool subprocess environment, but this is not a filesystem or
network sandbox.

Never store an API key in this directory. Rotate any key that has previously been
written to a source or log file.

## Tests

```bash
python3 -m unittest discover -s tests -v
```

Tests use a fake model client and do not spend API credits. QA integration tests
create local temporary presentations and use the installed system tools.
