# PPT Harness

A deliberately small Python agent loop that asks DeepSeek to build an editable
PowerPoint presentation. It loads the bundled complete official Anthropic PPTX skill
from `./pptx/SKILL.md`, exposes four local tools (`read_file`, `write_file`, `edit_file`,
and `bash`), and performs deterministic QA before accepting the result.

DeepSeek V4 is text-only. The harness renders slides and creates a contact sheet for
human review, but it does not send images to the model or pretend that the model
performed visual inspection.

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
