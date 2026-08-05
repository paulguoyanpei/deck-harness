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
├── deck.layout.json   (only if the model waived a deliberate overlap)
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
one element covering another, text escaping its container, text too wide for its
box, elements that should share a centreline but don't, and detached named flow
arrows. It reads shape coordinates and font metrics — nothing is rendered, so the
model runs it while it builds rather than only at the end. Cardinal arrows named
under `flow/` must touch another node at both endpoints.

```bash
python3 pptx/scripts/layout_check.py runs/demo/deck.pptx
```

It is advisory: it always exits 0 and never blocks QA. Nothing forces the model to
run it, so `grep layout_check runs/<name>/run.jsonl` is how you check that it did.

### Naming

The check is far sharper when the generator names what it creates, via pptxgenjs
`objectName`:

```js
s.addShape("roundRect", { ..., objectName: "card1" });        // a container
s.addText("标题",        { ..., objectName: "card1/title" });  // sits on card1
s.addShape("ellipse",   { ..., objectName: "@bg/blob" });     // background art
```

Paths declare intent: an element may overlap its own ancestors, and `@bg/` art may
be covered by anything, but two unrelated elements may not overlap. Unnamed
elements fall back to geometric inference — it still finds real defects, just with
more noise, and their warnings cannot be waived. Every report prints the naming
coverage per slide, because a clean result means much less on an unnamed deck.

### Waivers

Deliberate overlaps go in `deck.layout.json`, in the form the report prints:

```json
{"waivers": [{"slide": 5, "a": "chart/bar3", "b": "chart/callout",
              "ratio": 0.22, "reason": "callout deliberately marks that bar"}]}
```

A waiver needs a reason, needs both elements named, and holds only while the
overlap stays within `ratio` + 0.10 — move the element and it must be confirmed
again. Overlaps at or above 60% are never waivable.

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
