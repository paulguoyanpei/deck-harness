#!/usr/bin/env python3
"""Make one tool-free Messages API call using a bundled PPTX skill."""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parent
DEFAULT_SYSTEM = ROOT / "system_prompt.md"
DEFAULT_SKILL = ROOT / "pptx-api" / "SKILL.md"
DEFAULT_BASE_URL = "https://api.deepseek.com/anthropic"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_MAX_TOKENS = 128_000
SECRET_ENV_KEYS = {
    "DEEPSEEK_API_KEY",
    "ANTHROPIC_API_KEY",
    "ANTHROPIC_AUTH_TOKEN",
}


def load_prompt(system_path: Path, skill_path: Path) -> str:
    system = system_path.resolve().read_text(encoding="utf-8").strip()
    skill = skill_path.resolve().read_text(encoding="utf-8").strip()
    return f"{system}\n\n<one_shot_pptx_skill>\n{skill}\n</one_shot_pptx_skill>"


def response_data(response: Any) -> dict[str, Any]:
    if hasattr(response, "model_dump"):
        return response.model_dump(exclude_none=True)
    if isinstance(response, dict):
        return response
    raise TypeError(f"unsupported response type: {type(response).__name__}")


def response_text(data: dict[str, Any]) -> str:
    return "\n".join(
        block.get("text", "")
        for block in data.get("content", [])
        if block.get("type") == "text"
    ).strip()


def extract_javascript(text: str) -> str | None:
    matches = re.findall(r"```(?:javascript|js)\s*\n(.*?)```", text, re.DOTALL | re.IGNORECASE)
    return matches[0].strip() + "\n" if len(matches) == 1 else None


def _write_compile_report(out_dir: Path, report: dict[str, Any]) -> None:
    (out_dir / "compile.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def compile_javascript(source: str, out_dir: Path, timeout_seconds: int = 180) -> Path:
    """Execute extracted model code once; this is compilation, not an agent turn."""
    source_path = out_dir / "deck.js"
    deck_path = out_dir / "deck.pptx"
    source_path.write_text(source, encoding="utf-8")
    deck_path.unlink(missing_ok=True)
    env = os.environ.copy()
    for key in SECRET_ENV_KEYS:
        env.pop(key, None)
    node_modules = [ROOT / "node_modules", ROOT.parent / "ppt-harness" / "node_modules"]
    existing_node_path = env.get("NODE_PATH")
    env["NODE_PATH"] = os.pathsep.join(
        [str(path) for path in node_modules if path.is_dir()]
        + ([existing_node_path] if existing_node_path else [])
    )

    started = time.monotonic()
    try:
        completed = subprocess.run(
            ["node", source_path.name],
            cwd=out_dir,
            env=env,
            text=True,
            capture_output=True,
            timeout=timeout_seconds,
            check=False,
        )
        report = {
            "status": "passed" if completed.returncode == 0 and deck_path.is_file() else "failed",
            "command": ["node", source_path.name],
            "exit_code": completed.returncode,
            "stdout": completed.stdout,
            "stderr": completed.stderr,
            "duration_ms": round((time.monotonic() - started) * 1000),
            "deck_created": deck_path.is_file(),
        }
    except subprocess.TimeoutExpired as exc:
        report = {
            "status": "failed",
            "command": ["node", source_path.name],
            "exit_code": 124,
            "stdout": exc.stdout or "",
            "stderr": exc.stderr or "",
            "duration_ms": round((time.monotonic() - started) * 1000),
            "deck_created": deck_path.is_file(),
            "timed_out": True,
        }
    except OSError as exc:
        report = {
            "status": "failed",
            "command": ["node", source_path.name],
            "exit_code": 127,
            "stdout": "",
            "stderr": str(exc),
            "duration_ms": round((time.monotonic() - started) * 1000),
            "deck_created": False,
        }

    _write_compile_report(out_dir, report)
    if report["status"] != "passed":
        detail = report["stderr"].strip() or "deck.pptx was not created"
        raise RuntimeError(f"generated JavaScript compilation failed: {detail}")
    return deck_path


def generate(
    client: Any,
    request: str,
    out_dir: Path,
    *,
    model: str = DEFAULT_MODEL,
    max_tokens: int = DEFAULT_MAX_TOKENS,
    system_path: Path = DEFAULT_SYSTEM,
    skill_path: Path = DEFAULT_SKILL,
    auto_compile: bool = True,
    compile_timeout: int = 180,
) -> str:
    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    system = load_prompt(system_path, skill_path)

    # Deliberately one request: no tools, tool results, message replay, or QA loop.
    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": request}],
    )
    data = response_data(response)
    text = response_text(data)
    (out_dir / "response.md").write_text(text + ("\n" if text else ""), encoding="utf-8")
    (out_dir / "response.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    # Never let artifacts from an older run masquerade as this response's output.
    for artifact in ("deck.js", "deck.pptx", "compile.json"):
        (out_dir / artifact).unlink(missing_ok=True)
    if not auto_compile:
        _write_compile_report(out_dir, {"status": "disabled", "reason": "--no-compile"})
        return text
    if data.get("stop_reason") == "max_tokens":
        _write_compile_report(
            out_dir,
            {"status": "skipped", "reason": "API response was truncated at max_tokens"},
        )
        raise RuntimeError("API response was truncated at max_tokens; JavaScript was not executed")
    source = extract_javascript(text)
    if source is None:
        _write_compile_report(
            out_dir,
            {"status": "skipped", "reason": "expected exactly one complete JavaScript code block"},
        )
        raise RuntimeError("could not extract exactly one complete JavaScript code block")
    compile_javascript(source, out_dir, compile_timeout)
    return text


def create_client(api_key: str, base_url: str) -> Any:
    try:
        from anthropic import Anthropic
    except ImportError as exc:
        raise RuntimeError("missing dependency: install requirements.txt") from exc
    return Anthropic(api_key=api_key, base_url=base_url, timeout=600.0, max_retries=3)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="One-shot, tool-free PPT source generation API")
    parser.add_argument("request", nargs="*", help="presentation brief; stdin when omitted")
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--model", default=os.getenv("DEEPSEEK_MODEL", DEFAULT_MODEL))
    parser.add_argument("--base-url", default=os.getenv("DEEPSEEK_BASE_URL", DEFAULT_BASE_URL))
    parser.add_argument(
        "--max-tokens",
        type=int,
        default=int(os.getenv("DEEPSEEK_MAX_TOKENS", DEFAULT_MAX_TOKENS)),
        help=f"maximum output tokens (default: {DEFAULT_MAX_TOKENS}; DeepSeek V4 maximum: 384000)",
    )
    parser.add_argument("--system", type=Path, default=DEFAULT_SYSTEM)
    parser.add_argument("--skill", type=Path, default=DEFAULT_SKILL)
    parser.add_argument("--no-compile", action="store_true", help="save the API response without executing it")
    parser.add_argument("--compile-timeout", type=int, default=180, help="Node execution timeout in seconds")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    request = " ".join(args.request).strip()
    if not request and not sys.stdin.isatty():
        request = sys.stdin.read().strip()
    if not request:
        print("error: provide a presentation brief", file=sys.stderr)
        return 2
    if not 1 <= args.max_tokens <= 384_000:
        print("error: --max-tokens must be between 1 and 384000", file=sys.stderr)
        return 2
    if args.compile_timeout < 1:
        print("error: --compile-timeout must be positive", file=sys.stderr)
        return 2
    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        print("error: DEEPSEEK_API_KEY is not set", file=sys.stderr)
        return 2

    try:
        text = generate(
            create_client(api_key, args.base_url),
            request,
            args.out,
            model=args.model,
            max_tokens=args.max_tokens,
            system_path=args.system,
            skill_path=args.skill,
            auto_compile=not args.no_compile,
            compile_timeout=args.compile_timeout,
        )
    except (OSError, RuntimeError, TypeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    print(text)
    print(f"\nSaved API response to {args.out.resolve()}")
    if not args.no_compile:
        print(f"Generated PPTX: {args.out.resolve() / 'deck.pptx'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
