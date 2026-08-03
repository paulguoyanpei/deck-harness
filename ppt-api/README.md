# ppt-api

A one-shot control experiment for PPT generation. It sends one user request with
`system_prompt.md` and the bundled `pptx-api/SKILL.md` to DeepSeek's Anthropic-
compatible endpoint.

It declares no tools and runs no agent loop. After the single response returns,
the CLI extracts exactly one complete JavaScript code block, saves `deck.js`, and
runs it once with Node to create `deck.pptx`. This compile step is deterministic;
it never sends errors back to the model and performs no repair loop or visual QA.

```bash
cd /home/gyp/repo/api/ppt-api
python3 -m pip install -r requirements.txt
npm install
export DEEPSEEK_API_KEY="your-key"
python3 ppt_api.py --out ./runs/demo "制作一份10页的人工智能科普演示文稿"
```

The default output budget is 128,000 tokens. Override it per run with
`--max-tokens` or `DEEPSEEK_MAX_TOKENS`; DeepSeek V4 currently allows up to
384,000 output tokens.

Outputs:

- `response.md`: model-facing text exactly as returned
- `response.json`: full API response, excluding credentials
- `deck.js`: extracted JavaScript source
- `deck.pptx`: output created by running `deck.js`
- `compile.json`: Node exit code, stdout, stderr, duration, and compile status

Compilation is skipped and the CLI exits with an error when the API reports
`stop_reason=max_tokens` or the response does not contain exactly one complete
JavaScript block. Use `--no-compile` to retain the previous response-only behavior.

The generated JavaScript is untrusted model output and is executed with the current
user's filesystem permissions. API credential environment variables are removed
from the Node subprocess, but run this experiment only in a disposable output
directory.
