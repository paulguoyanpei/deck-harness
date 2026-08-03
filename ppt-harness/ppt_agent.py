#!/usr/bin/env python3
"""A small, transparent PPT-building agent loop for DeepSeek."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_BASE_URL = "https://api.deepseek.com/anthropic"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_MAX_STEPS = 40
DEFAULT_SKILL_PATH = Path(__file__).resolve().parent / "pptx" / "SKILL.md"
MAX_TOOL_OUTPUT = 30_000
PLACEHOLDER_RE = re.compile(r"\b(?:TODO|lorem|ipsum)\b|\[insert", re.IGNORECASE)
SECRET_ENV_KEYS = {
    "DEEPSEEK_API_KEY",
    "ANTHROPIC_API_KEY",
    "ANTHROPIC_AUTH_TOKEN",
    "ANTHROPIC_API_KEY",
}

SYSTEM_HEADER = """\
You are an autonomous presentation agent. The complete official Anthropic PPTX
skill below is your authoritative instruction set for presentation creation,
editing, analysis, design, and QA.
"""

RUNTIME_CONTRACT = """\
Work in WORKING_DIR and produce WORKING_DIR/deck.pptx with its editable generator
source. Resolve skill-relative scripts against PPTX_SKILL_DIR. Direct image
inspection is unavailable; use programmatic QA where needed.
"""

TOOLS = [
    {
        "name": "read_file",
        "description": "Read a UTF-8 text file in the working directory with line numbers.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "offset": {"type": "integer", "minimum": 0},
                "limit": {"type": "integer", "minimum": 1, "maximum": 2000},
            },
            "required": ["path"],
            "additionalProperties": False,
        },
    },
    {
        "name": "write_file",
        "description": "Create or replace a UTF-8 text file in the working directory.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "content": {"type": "string"},
            },
            "required": ["path", "content"],
            "additionalProperties": False,
        },
    },
    {
        "name": "edit_file",
        "description": "Replace one exact, unique text fragment in a UTF-8 file.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "old_text": {"type": "string"},
                "new_text": {"type": "string"},
                "replace_all": {"type": "boolean"},
            },
            "required": ["path", "old_text", "new_text"],
            "additionalProperties": False,
        },
    },
    {
        "name": "bash",
        "description": "Run a Bash command in the working directory.",
        "input_schema": {
            "type": "object",
            "properties": {
                "command": {"type": "string"},
                "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 900},
            },
            "required": ["command"],
            "additionalProperties": False,
        },
    },
]


def clip(text: str, limit: int = MAX_TOOL_OUTPUT) -> str:
    if len(text) <= limit:
        return text
    half = limit // 2
    omitted = len(text) - limit
    return f"{text[:half]}\n... <{omitted} chars omitted> ...\n{text[-half:]}"


def load_system_prompt(
    skill_path: Path = DEFAULT_SKILL_PATH,
    workdir: Path | None = None,
) -> tuple[str, Path, str]:
    skill_path = skill_path.resolve()
    if not skill_path.is_file():
        raise RuntimeError(f"PPTX skill not found: {skill_path}")
    skill_text = skill_path.read_text(encoding="utf-8")
    digest = hashlib.sha256(skill_text.encode("utf-8")).hexdigest()
    runtime = RUNTIME_CONTRACT.replace("PPTX_SKILL_DIR", str(skill_path.parent)).replace(
        "WORKING_DIR", str((workdir or Path.cwd()).resolve())
    )
    prompt = (
        SYSTEM_HEADER
        + "\n<official_pptx_skill>\n"
        + skill_text
        + "\n</official_pptx_skill>\n\n"
        + runtime
    )
    return prompt, skill_path.parent, digest


def jsonable(value: Any) -> Any:
    if hasattr(value, "model_dump"):
        return value.model_dump(exclude_none=True)
    if isinstance(value, list):
        return [jsonable(item) for item in value]
    if isinstance(value, dict):
        return {key: jsonable(item) for key, item in value.items()}
    return value


def field(value: Any, name: str, default: Any = None) -> Any:
    if isinstance(value, dict):
        return value.get(name, default)
    return getattr(value, name, default)


@dataclass
class CommandResult:
    command: list[str] | str
    exit_code: int
    stdout: str
    stderr: str
    duration_ms: int
    timed_out: bool = False

    def model_content(self) -> str:
        """Return a concise Claude Code-style observation for the model."""
        output = "\n".join(part.rstrip() for part in (self.stdout, self.stderr) if part.strip())
        if self.timed_out:
            prefix = "Command timed out (exit code 124)"
        elif self.exit_code != 0:
            prefix = f"Exit code {self.exit_code}"
        elif not output:
            return "Command completed successfully."
        else:
            return clip(output)
        return clip(f"{prefix}\n{output}" if output else prefix)

    def log_data(self) -> dict[str, Any]:
        return {
            "command": self.command,
            "exit_code": self.exit_code,
            "stdout": self.stdout,
            "stderr": self.stderr,
            "duration_ms": self.duration_ms,
            "timed_out": self.timed_out,
        }


@dataclass
class ToolExecution:
    """Separate the model-facing observation from diagnostic log data."""

    model_content: str
    is_error: bool
    details: dict[str, Any]

    def __iter__(self):
        """Keep the original ``result, is_error = execute(...)`` API usable."""
        yield self.model_content
        yield self.is_error


def run_command(
    command: list[str] | str,
    cwd: Path,
    timeout: int = 180,
    *,
    shell: bool = False,
    strip_secrets: bool = False,
    max_output: int | None = MAX_TOOL_OUTPUT,
    env_overrides: dict[str, str] | None = None,
) -> CommandResult:
    env = os.environ.copy()
    if strip_secrets:
        for key in SECRET_ENV_KEYS:
            env.pop(key, None)
    if env_overrides:
        env.update(env_overrides)
    started = time.monotonic()
    try:
        proc = subprocess.run(
            command,
            cwd=cwd,
            env=env,
            shell=shell,
            executable="/bin/bash" if shell else None,
            text=True,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        return CommandResult(
            command=command,
            exit_code=proc.returncode,
            stdout=clip(proc.stdout, max_output) if max_output is not None else proc.stdout,
            stderr=clip(proc.stderr, max_output) if max_output is not None else proc.stderr,
            duration_ms=round((time.monotonic() - started) * 1000),
        )
    except subprocess.TimeoutExpired as exc:
        stdout = exc.stdout.decode() if isinstance(exc.stdout, bytes) else (exc.stdout or "")
        stderr = exc.stderr.decode() if isinstance(exc.stderr, bytes) else (exc.stderr or "")
        return CommandResult(
            command=command,
            exit_code=124,
            stdout=clip(stdout),
            stderr=clip(stderr),
            duration_ms=round((time.monotonic() - started) * 1000),
            timed_out=True,
        )
    except OSError as exc:
        return CommandResult(
            command=command,
            exit_code=127,
            stdout="",
            stderr=str(exc),
            duration_ms=round((time.monotonic() - started) * 1000),
        )


class ToolRunner:
    def __init__(self, workdir: Path):
        self.workdir = workdir.resolve()

    def _path(self, raw: str) -> Path:
        candidate = (self.workdir / raw).resolve()
        try:
            candidate.relative_to(self.workdir)
        except ValueError as exc:
            raise ValueError("path must stay inside the working directory") from exc
        return candidate

    def execute(self, name: str, payload: dict[str, Any]) -> ToolExecution:
        try:
            if name == "read_file":
                content = self.read_file(**payload)
                return ToolExecution(content, False, {"path": payload.get("path")})
            if name == "write_file":
                content = self.write_file(**payload)
                return ToolExecution(
                    content,
                    False,
                    {
                        "path": str(self._path(payload["path"])),
                        "characters_written": len(payload["content"]),
                    },
                )
            if name == "edit_file":
                content, replacements = self.edit_file(**payload)
                return ToolExecution(
                    content,
                    False,
                    {"path": str(self._path(payload["path"])), "replacements": replacements},
                )
            if name == "bash":
                result = self.bash(**payload)
                return ToolExecution(result.model_content(), result.exit_code != 0, result.log_data())
            message = f"Unknown tool: {name}"
            return ToolExecution(message, True, {"error": message})
        except (OSError, UnicodeError, ValueError, TypeError) as exc:
            message = f"{type(exc).__name__}: {exc}"
            return ToolExecution(message, True, {"error": message})

    def read_file(self, path: str, offset: int = 0, limit: int = 400) -> str:
        target = self._path(path)
        lines = target.read_text(encoding="utf-8").splitlines()
        selected = lines[offset : offset + limit]
        body = "\n".join(f"{offset + index + 1:>6}  {line}" for index, line in enumerate(selected))
        return clip(f"path={path} lines={len(lines)} showing={offset + 1}-{offset + len(selected)}\n{body}")

    def write_file(self, path: str, content: str) -> str:
        target = self._path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
        return (
            f"File created successfully at: {target}\n"
            "File state is current in your context; no need to read it back."
        )

    def edit_file(
        self,
        path: str,
        old_text: str,
        new_text: str,
        replace_all: bool = False,
    ) -> tuple[str, int]:
        if not old_text:
            raise ValueError("old_text must not be empty")
        target = self._path(path)
        content = target.read_text(encoding="utf-8")
        matches = content.count(old_text)
        if matches == 0:
            raise ValueError("old_text was not found")
        if not replace_all and matches != 1:
            raise ValueError(f"old_text matched {matches} times; provide a unique fragment or replace_all=true")
        updated = content.replace(old_text, new_text, -1 if replace_all else 1)
        target.write_text(updated, encoding="utf-8")
        replacements = matches if replace_all else 1
        return (
            f"File updated successfully at: {target}\n"
            "File state is current in your context; no need to read it back.",
            replacements,
        )

    def bash(self, command: str, timeout_seconds: int = 180) -> CommandResult:
        return run_command(
            command,
            self.workdir,
            timeout_seconds,
            shell=True,
            strip_secrets=True,
        )


class JsonlLogger:
    def __init__(self, path: Path):
        self.path = path

    def write(self, kind: str, data: Any) -> None:
        record = {"timestamp": time.time(), "kind": kind, "data": jsonable(data)}
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def _pdf_bounds(pdf: Path, qa_dir: Path, workdir: Path) -> tuple[bool, str]:
    bbox = qa_dir / "bbox.html"
    result = run_command(["pdftotext", "-bbox", str(pdf), str(bbox)], workdir, 120)
    if result.exit_code != 0 or not bbox.is_file():
        return False, result.stderr or "pdftotext did not create bbox.html"
    html = bbox.read_text(encoding="utf-8", errors="replace")
    pages = re.findall(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', html, re.DOTALL)
    if not pages:
        return False, "no pages found in bbox output"
    issues: list[str] = []
    word_re = re.compile(
        r'<word xMin="([\d.-]+)" yMin="([\d.-]+)" '
        r'xMax="([\d.-]+)" yMax="([\d.-]+)"[^>]*>(.*?)</word>',
        re.DOTALL,
    )
    for page_number, (width, height, body) in enumerate(pages, 1):
        page_width, page_height = float(width), float(height)
        for x1, y1, x2, y2, text in word_re.findall(body):
            coords = tuple(map(float, (x1, y1, x2, y2)))
            if coords[0] < -0.5 or coords[1] < -0.5 or coords[2] > page_width + 0.5 or coords[3] > page_height + 0.5:
                plain = re.sub(r"<[^>]+>", "", text).strip()
                issues.append(f"slide {page_number}: {plain!r} at {coords}")
                if len(issues) == 10:
                    break
        if len(issues) == 10:
            break
    return not issues, f"pages={len(pages)}; " + ("no text outside page bounds" if not issues else "; ".join(issues))


def _publish_qa_artifacts(qa_dir: Path, workdir: Path) -> None:
    os.replace(qa_dir / "deck.pdf", workdir / "deck.pdf")
    os.replace(qa_dir / "preview.jpg", workdir / "preview.jpg")
    next_slides = workdir / ".slides-next"
    old_slides = workdir / ".slides-old"
    shutil.rmtree(next_slides, ignore_errors=True)
    shutil.rmtree(old_slides, ignore_errors=True)
    shutil.copytree(qa_dir / "slides", next_slides)
    slides = workdir / "slides"
    if slides.exists():
        slides.rename(old_slides)
    next_slides.rename(slides)
    shutil.rmtree(old_slides, ignore_errors=True)


def qa_presentation(workdir: Path, skill_dir: Path = DEFAULT_SKILL_PATH.parent) -> tuple[bool, str]:
    workdir = workdir.expanduser().resolve()
    pptx = workdir / "deck.pptx"
    report_path = workdir / "qa.txt"
    skill_dir = skill_dir.resolve()
    checks: list[tuple[str, bool, str]] = []

    def add(name: str, ok: bool, detail: str) -> None:
        checks.append((name, ok, detail))

    add("pptx_exists", pptx.is_file(), "deck.pptx exists")
    if not pptx.is_file():
        report = _format_qa(checks)
        report_path.write_text(report, encoding="utf-8")
        return False, report

    unzip_result = run_command(["unzip", "-t", str(pptx)], workdir, 120)
    add("pptx_zip", unzip_result.exit_code == 0, unzip_result.stderr or unzip_result.stdout[-1000:])

    validate_script = skill_dir / "scripts" / "office" / "validate.py"
    validate = run_command([sys.executable, str(validate_script), str(pptx)], workdir, 180)
    add("official_validation", validate.exit_code == 0, validate.stderr or validate.stdout)

    slide_count = 0
    try:
        with zipfile.ZipFile(pptx) as archive:
            slide_count = sum(
                bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", name))
                for name in archive.namelist()
            )
        add("pptx_slides", slide_count > 0, f"slides={slide_count}")
    except zipfile.BadZipFile as exc:
        add("pptx_slides", False, str(exc))

    with tempfile.TemporaryDirectory(prefix=".qa-", dir=workdir) as qa_temp:
        qa_dir = Path(qa_temp).resolve()
        pdf = qa_dir / "deck.pdf"
        slides_dir = qa_dir / "slides"
        preview = qa_dir / "preview.jpg"
        slides_dir.mkdir()
        xdg_runtime = qa_dir / "xdg-runtime"
        xdg_config = qa_dir / "xdg-config"
        xdg_runtime.mkdir(mode=0o700)
        xdg_config.mkdir()
        soffice_script = skill_dir / "scripts" / "office" / "soffice.py"
        convert = run_command(
            [sys.executable, str(soffice_script), "--headless", "--convert-to", "pdf", "--outdir", str(qa_dir), str(pptx)],
            workdir,
            180,
            env_overrides={
                "XDG_RUNTIME_DIR": str(xdg_runtime),
                "XDG_CONFIG_HOME": str(xdg_config),
                "GSETTINGS_BACKEND": "memory",
            },
        )
        add("pdf_conversion", convert.exit_code == 0 and pdf.is_file(), convert.stderr or convert.stdout)

        pdf_pages = 0
        if pdf.is_file():
            info = run_command(["pdfinfo", str(pdf)], workdir, 60)
            match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
            pdf_pages = int(match.group(1)) if match else 0
            add("pdf_pages", pdf_pages == slide_count and pdf_pages > 0, f"pptx={slide_count} pdf={pdf_pages}")
            text = run_command(["markitdown", str(pptx)], workdir, 180, max_output=None)
            placeholders = sorted(set(found.group(0) for found in PLACEHOLDER_RE.finditer(text.stdout)))
            add("placeholders", text.exit_code == 0 and not placeholders, f"found={placeholders}" if placeholders else "none found")
            bounds_ok, bounds_detail = _pdf_bounds(pdf, qa_dir, workdir)
            add("text_bounds", bounds_ok, bounds_detail)

            render = run_command(["pdftoppm", "-jpeg", "-r", "150", str(pdf), str(slides_dir / "slide")], workdir, 180)
            jpgs = sorted(slides_dir.glob("slide-*.jpg"))
            add("render", render.exit_code == 0 and len(jpgs) == pdf_pages, f"jpgs={len(jpgs)} pages={pdf_pages}")
            if jpgs:
                montage = run_command(
                    ["montage", *map(str, jpgs), "-thumbnail", "360x203", "-tile", "4x", "-geometry", "+8+8", str(preview)],
                    workdir,
                    180,
                )
                add("preview", montage.exit_code == 0 and preview.is_file(), montage.stderr or "preview.jpg created")

        passed = all(ok for _, ok, _ in checks)
        if passed:
            _publish_qa_artifacts(qa_dir, workdir)

    report = _format_qa(checks)
    report_path.write_text(report, encoding="utf-8")
    return passed, report


def _format_qa(checks: list[tuple[str, bool, str]]) -> str:
    lines = ["PASS" if all(ok for _, ok, _ in checks) else "FAIL"]
    for name, ok, detail in checks:
        clean_detail = " ".join(detail.split())
        lines.append(f"[{'PASS' if ok else 'FAIL'}] {name}: {clean_detail}")
    return "\n".join(lines) + "\n"


def create_client(api_key: str, base_url: str) -> Any:
    try:
        from anthropic import Anthropic
    except ImportError as exc:
        raise RuntimeError("missing dependency: run `python3 -m pip install -r requirements.txt`") from exc
    return Anthropic(api_key=api_key, base_url=base_url, timeout=600.0, max_retries=3)


def run_agent(
    client: Any,
    request: str,
    workdir: Path,
    *,
    model: str = DEFAULT_MODEL,
    max_steps: int = DEFAULT_MAX_STEPS,
    skill_path: Path = DEFAULT_SKILL_PATH,
) -> str:
    workdir = workdir.expanduser().resolve()
    workdir.mkdir(parents=True, exist_ok=True)
    runner = ToolRunner(workdir)
    logger = JsonlLogger(workdir / "run.jsonl")
    system_prompt, skill_dir, skill_digest = load_system_prompt(skill_path, workdir)
    logger.write("skill", {"path": str(skill_path.resolve()), "sha256": skill_digest})
    messages: list[dict[str, Any]] = [{"role": "user", "content": request}]

    for step in range(1, max_steps + 1):
        print(f"\n[{step}/{max_steps}] asking {model}...", flush=True)
        logger.write("request", {"step": step, "model": model, "message_count": len(messages)})
        response = client.messages.create(
            model=model,
            max_tokens=32_000,
            system=system_prompt,
            messages=messages,
            tools=TOOLS,
        )
        content = jsonable(field(response, "content", []))
        logger.write("response", {"step": step, "stop_reason": field(response, "stop_reason"), "content": content})
        messages.append({"role": "assistant", "content": content})

        tool_blocks = [block for block in content if field(block, "type") == "tool_use"]
        for block in content:
            if field(block, "type") == "text" and field(block, "text"):
                print(field(block, "text"))

        if tool_blocks:
            results = []
            for block in tool_blocks:
                name = field(block, "name")
                payload = field(block, "input", {})
                print(f"  → {name}", flush=True)
                execution = runner.execute(name, payload)
                logger.write(
                    "tool",
                    {
                        "name": name,
                        "input": payload,
                        "model_content": execution.model_content,
                        "details": execution.details,
                        "is_error": execution.is_error,
                    },
                )
                print(f"    {clip(execution.model_content, 500)}", flush=True)
                results.append(
                    {
                        "type": "tool_result",
                        "tool_use_id": field(block, "id"),
                        "content": execution.model_content,
                        "is_error": execution.is_error,
                    }
                )
            messages.append({"role": "user", "content": results})
            continue

        print("  → deterministic QA", flush=True)
        passed, report = qa_presentation(workdir, skill_dir)
        logger.write("qa", {"passed": passed, "report": report})
        print(report, end="")
        if passed:
            final_text = "\n".join(
                field(block, "text", "") for block in content if field(block, "type") == "text"
            ).strip()
            return final_text or "Presentation generated and QA passed."

        messages.append(
            {
                "role": "user",
                "content": "Deterministic QA failed. Repair the deck and try again.\n\n" + report,
            }
        )

    raise RuntimeError(f"agent exceeded the maximum of {max_steps} steps")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build an editable PPTX with a small DeepSeek agent loop.")
    parser.add_argument("request", nargs="*", help="presentation request; read from stdin when omitted")
    parser.add_argument("--out", required=True, type=Path, help="output directory for this run")
    parser.add_argument("--model", default=os.getenv("DEEPSEEK_MODEL", DEFAULT_MODEL))
    parser.add_argument("--max-steps", type=int, default=DEFAULT_MAX_STEPS)
    parser.add_argument("--skill", type=Path, default=DEFAULT_SKILL_PATH, help="path to the complete PPTX SKILL.md")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    request = " ".join(args.request).strip()
    if not request and not sys.stdin.isatty():
        request = sys.stdin.read().strip()
    if not request:
        print("error: provide a presentation request as arguments or stdin", file=sys.stderr)
        return 2
    if args.max_steps < 1:
        print("error: --max-steps must be at least 1", file=sys.stderr)
        return 2
    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        print("error: DEEPSEEK_API_KEY is not set", file=sys.stderr)
        return 2

    try:
        client = create_client(api_key, os.getenv("DEEPSEEK_BASE_URL", DEFAULT_BASE_URL))
        final = run_agent(
            client,
            request,
            args.out,
            model=args.model,
            max_steps=args.max_steps,
            skill_path=args.skill,
        )
    except (RuntimeError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    print(f"\nDone: {args.out.resolve() / 'deck.pptx'}")
    print(final)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
