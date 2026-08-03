#!/usr/bin/env python3
"""Send a PPT prompt to the Claude Code CLI running on DeepSeek V4 Flash.

    python3 cc_agent.py '做一份 10 页的向量数据库科普演示文稿'

Print mode plus bypassed permissions plus a closed stdin: the CLI answers and
exits, and nothing in the run can stop to ask a follow-up.
"""

from __future__ import annotations

import os
import subprocess
import sys


BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/anthropic")
MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-v4-flash")

# Claude Code picks a model per tier; any tier left unset sends a claude-* name
# upstream and 404s. The last one covers Task subagents.
MODEL_ENV_KEYS = (
    "ANTHROPIC_MODEL",
    "ANTHROPIC_DEFAULT_OPUS_MODEL",
    "ANTHROPIC_DEFAULT_SONNET_MODEL",
    "ANTHROPIC_DEFAULT_HAIKU_MODEL",
    "CLAUDE_CODE_SUBAGENT_MODEL",
)


def main(argv: list[str]) -> int:
    prompt = " ".join(argv[1:]).strip()
    if not prompt:
        print(f"usage: {argv[0]} '<presentation prompt>'", file=sys.stderr)
        return 2
    api_key = os.getenv("DEEPSEEK_API_KEY") or os.getenv("ANTHROPIC_AUTH_TOKEN")
    if not api_key:
        print("error: set DEEPSEEK_API_KEY", file=sys.stderr)
        return 2

    env = os.environ.copy()
    env.pop("ANTHROPIC_API_KEY", None)  # would take precedence over the DeepSeek token
    env["ANTHROPIC_BASE_URL"] = BASE_URL
    env["ANTHROPIC_AUTH_TOKEN"] = api_key
    for key in MODEL_ENV_KEYS:
        env[key] = MODEL
    # DeepSeek's compatibility layer reads output_config.effort, not budget_tokens.
    env["CLAUDE_CODE_EFFORT_LEVEL"] = os.getenv("CLAUDE_CODE_EFFORT_LEVEL", "high")

    command = [
        "claude",
        "--print",  # non-interactive: answer and exit instead of opening a REPL
        "--permission-mode",
        "bypassPermissions",  # no tool-approval prompt can appear
        "--model",
        MODEL,
        prompt,
    ]
    return subprocess.run(command, env=env, stdin=subprocess.DEVNULL).returncode


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
