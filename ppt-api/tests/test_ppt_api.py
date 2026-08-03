from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import ppt_api  # noqa: E402


class FakeMessages:
    def __init__(self, response):
        self.response = response
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        return self.response


class FakeClient:
    def __init__(self, response):
        self.messages = FakeMessages(response)


class PptApiTests(unittest.TestCase):
    def test_default_output_budget_is_128k(self):
        args = ppt_api.parse_args(["--out", "runs/demo", "Build"])
        self.assertEqual(args.max_tokens, 128_000)

    def test_cli_accepts_deepseek_maximum(self):
        args = ppt_api.parse_args(
            ["--out", "runs/demo", "--max-tokens", "384000", "Build"]
        )
        self.assertEqual(args.max_tokens, 384_000)

    def test_exactly_one_call_without_tools_and_extracts_source(self):
        response = {
            "id": "msg-1",
            "stop_reason": "end_turn",
            "content": [{"type": "text", "text": "```javascript\nconsole.log('deck');\n```"}],
        }
        client = FakeClient(response)
        with tempfile.TemporaryDirectory() as temp, patch.object(
            ppt_api, "compile_javascript", return_value=Path(temp) / "deck.pptx"
        ) as compile_js:
            root = Path(temp)
            text = ppt_api.generate(client, "Build ten slides", root)
            saved_json = json.loads((root / "response.json").read_text(encoding="utf-8"))

        self.assertIn("console.log", text)
        self.assertEqual(len(client.messages.calls), 1)
        call = client.messages.calls[0]
        self.assertNotIn("tools", call)
        self.assertEqual(call["messages"], [{"role": "user", "content": "Build ten slides"}])
        self.assertIn("<one_shot_pptx_skill>", call["system"])
        self.assertEqual(saved_json["id"], "msg-1")
        compile_js.assert_called_once_with("console.log('deck');\n", root, 180)

    def test_ambiguous_or_missing_code_is_not_compiled(self):
        response = {"content": [{"type": "text", "text": "No fenced source"}]}
        client = FakeClient(response)
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            with self.assertRaisesRegex(RuntimeError, "could not extract"):
                ppt_api.generate(client, "Build", root)
            self.assertFalse((root / "deck.js").exists())
            report = json.loads((root / "compile.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "skipped")

    def test_truncated_response_is_never_compiled(self):
        response = {
            "stop_reason": "max_tokens",
            "content": [{"type": "text", "text": "```javascript\npartial"}],
        }
        client = FakeClient(response)
        with tempfile.TemporaryDirectory() as temp, patch.object(
            ppt_api, "compile_javascript"
        ) as compile_js:
            root = Path(temp)
            (root / "deck.js").write_text("stale")
            (root / "deck.pptx").write_bytes(b"stale")
            with self.assertRaisesRegex(RuntimeError, "truncated"):
                ppt_api.generate(client, "Build", root)
            compile_js.assert_not_called()
            self.assertFalse((root / "deck.js").exists())
            self.assertFalse((root / "deck.pptx").exists())
            report = json.loads((root / "compile.json").read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "skipped")

    def test_compile_runs_node_without_api_key_and_requires_deck(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)

            def fake_run(*args, **kwargs):
                self.assertNotIn("DEEPSEEK_API_KEY", kwargs["env"])
                self.assertIn("node_modules", kwargs["env"]["NODE_PATH"])
                (root / "deck.pptx").write_bytes(b"pptx")
                return SimpleNamespace(returncode=0, stdout="made deck", stderr="")

            with patch.dict("os.environ", {"DEEPSEEK_API_KEY": "secret"}), patch.object(
                ppt_api.subprocess, "run", side_effect=fake_run
            ):
                result = ppt_api.compile_javascript("// source\n", root)

            self.assertEqual(result, root / "deck.pptx")
            self.assertEqual((root / "deck.js").read_text(), "// source\n")
            report = json.loads((root / "compile.json").read_text())
            self.assertEqual(report["status"], "passed")

    def test_missing_prompt_file_fails_before_api_call(self):
        client = FakeClient({"content": []})
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaises(FileNotFoundError):
                ppt_api.generate(
                    client,
                    "Build",
                    Path(temp) / "out",
                    system_path=Path(temp) / "missing.md",
                )
        self.assertEqual(client.messages.calls, [])


if __name__ == "__main__":
    unittest.main()
