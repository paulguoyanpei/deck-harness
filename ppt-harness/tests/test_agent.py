from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import ppt_agent  # noqa: E402


class FakeMessages:
    def __init__(self, responses):
        self.responses = iter(responses)
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(deepcopy(kwargs))
        return next(self.responses)


class FakeClient:
    def __init__(self, responses):
        self.messages = FakeMessages(responses)


class ToolRunnerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.runner = ppt_agent.ToolRunner(self.root)

    def tearDown(self):
        self.temp.cleanup()

    def test_write_read_and_unique_edit(self):
        result, error = self.runner.execute(
            "write_file", {"path": "src/deck.py", "content": "alpha\nbeta\n"}
        )
        self.assertFalse(error)
        self.assertIn("File created successfully at:", result)
        self.assertIn("no need to read it back", result)

        result, error = self.runner.execute(
            "read_file", {"path": "src/deck.py", "offset": 1, "limit": 1}
        )
        self.assertFalse(error)
        self.assertIn("     2  beta", result)

        result, error = self.runner.execute(
            "edit_file",
            {"path": "src/deck.py", "old_text": "beta", "new_text": "gamma"},
        )
        self.assertFalse(error)
        self.assertEqual((self.root / "src/deck.py").read_text(), "alpha\ngamma\n")
        self.assertIn("File updated successfully at:", result)

    def test_edit_rejects_ambiguous_match(self):
        (self.root / "deck.py").write_text("same same", encoding="utf-8")
        result, error = self.runner.execute(
            "edit_file",
            {"path": "deck.py", "old_text": "same", "new_text": "new"},
        )
        self.assertTrue(error)
        self.assertIn("matched 2 times", result)

    def test_file_tools_cannot_escape_workdir(self):
        result, error = self.runner.execute(
            "write_file", {"path": "../escape.txt", "content": "no"}
        )
        self.assertTrue(error)
        self.assertIn("inside the working directory", result)

    def test_bash_strips_api_keys(self):
        with patch.dict(os.environ, {"DEEPSEEK_API_KEY": "super-secret"}):
            result = self.runner.bash('test -z "$DEEPSEEK_API_KEY"')
        self.assertEqual(result.exit_code, 0)

    def test_bash_reports_failure_and_timeout(self):
        failed = self.runner.bash("echo bad >&2; exit 7")
        self.assertEqual(failed.exit_code, 7)
        self.assertIn("bad", failed.stderr)

        timed_out = self.runner.bash("sleep 2", timeout_seconds=1)
        self.assertTrue(timed_out.timed_out)
        self.assertEqual(timed_out.exit_code, 124)

    def test_bash_model_content_is_concise_but_logs_are_structured(self):
        execution = self.runner.execute("bash", {"command": "printf out; printf err >&2; exit 7"})
        self.assertTrue(execution.is_error)
        self.assertEqual(execution.model_content, "Exit code 7\nout\nerr")
        self.assertNotIn('"command"', execution.model_content)
        self.assertEqual(execution.details["command"], "printf out; printf err >&2; exit 7")
        self.assertEqual(execution.details["exit_code"], 7)
        self.assertEqual(execution.details["stdout"], "out")
        self.assertEqual(execution.details["stderr"], "err")

        empty_success = self.runner.execute("bash", {"command": "true"})
        self.assertEqual(empty_success.model_content, "Command completed successfully.")

    def test_tool_output_is_truncated(self):
        result = self.runner.bash("python3 -c 'print(\"x\" * 40000)'")
        self.assertLessEqual(len(result.stdout), ppt_agent.MAX_TOOL_OUTPUT + 100)
        self.assertIn("chars omitted", result.stdout)


class AgentLoopTests(unittest.TestCase):
    def test_tool_results_and_failed_qa_return_to_model(self):
        responses = [
            {
                "stop_reason": "tool_use",
                "content": [
                    {
                        "type": "tool_use",
                        "id": "tool-1",
                        "name": "write_file",
                        "input": {"path": "deck.py", "content": "print('one')\n"},
                    }
                ],
            },
            {"stop_reason": "end_turn", "content": [{"type": "text", "text": "Done"}]},
            {
                "stop_reason": "tool_use",
                "content": [
                    {
                        "type": "tool_use",
                        "id": "tool-2",
                        "name": "edit_file",
                        "input": {
                            "path": "deck.py",
                            "old_text": "one",
                            "new_text": "two",
                        },
                    }
                ],
            },
            {"stop_reason": "end_turn", "content": [{"type": "text", "text": "Fixed"}]},
        ]
        client = FakeClient(responses)
        with tempfile.TemporaryDirectory() as temp, patch.object(
            ppt_agent, "qa_presentation", side_effect=[(False, "FAIL\nmissing"), (True, "PASS\n")]
        ):
            final = ppt_agent.run_agent(client, "Build slides", Path(temp), max_steps=6)
            log_lines = (Path(temp) / "run.jsonl").read_text().splitlines()

        self.assertEqual(final, "Fixed")
        self.assertEqual(len(client.messages.calls), 4)
        official_skill = ppt_agent.DEFAULT_SKILL_PATH.read_text(encoding="utf-8")
        self.assertIn(official_skill, client.messages.calls[0]["system"])
        self.assertIn(str(ppt_agent.DEFAULT_SKILL_PATH.parent), client.messages.calls[0]["system"])
        self.assertIn(str(Path(temp).resolve()), client.messages.calls[0]["system"])
        self.assertNotIn("HARNESS_PROMPT", client.messages.calls[0]["system"])
        self.assertIn("use programmatic QA where needed", client.messages.calls[0]["system"])
        self.assertNotIn("Do not run validate.py", client.messages.calls[0]["system"])
        second_messages = client.messages.calls[1]["messages"]
        self.assertEqual(second_messages[-1]["content"][0]["tool_use_id"], "tool-1")
        self.assertIn("File created successfully at:", second_messages[-1]["content"][0]["content"])
        tool_log = next(json.loads(line) for line in log_lines if json.loads(line)["kind"] == "tool")
        self.assertIn("model_content", tool_log["data"])
        self.assertIn("details", tool_log["data"])
        third_messages = client.messages.calls[2]["messages"]
        self.assertIn("Deterministic QA failed", third_messages[-1]["content"])
        self.assertTrue(any(json.loads(line)["kind"] == "qa" for line in log_lines))

    def test_missing_skill_fails_before_api_call(self):
        client = FakeClient([])
        with tempfile.TemporaryDirectory() as temp:
            missing = Path(temp) / "missing-skill.md"
            with self.assertRaisesRegex(RuntimeError, "PPTX skill not found"):
                ppt_agent.run_agent(client, "Build", Path(temp) / "run", skill_path=missing)
        self.assertEqual(client.messages.calls, [])

    def test_max_steps_is_enforced(self):
        response = {
            "stop_reason": "tool_use",
            "content": [
                {
                    "type": "tool_use",
                    "id": "tool-1",
                    "name": "bash",
                    "input": {"command": "true"},
                }
            ],
        }
        client = FakeClient([response])
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaisesRegex(RuntimeError, "maximum of 1 steps"):
                ppt_agent.run_agent(client, "Build", Path(temp), max_steps=1)

    def test_relative_workdir_is_resolved_before_qa(self):
        client = FakeClient(
            [{"stop_reason": "end_turn", "content": [{"type": "text", "text": "Done"}]}]
        )
        with tempfile.TemporaryDirectory() as temp:
            parent = Path(temp)
            old_cwd = Path.cwd()
            try:
                os.chdir(parent)
                with patch.object(ppt_agent, "qa_presentation", return_value=(True, "PASS\n")) as qa:
                    ppt_agent.run_agent(client, "Build", Path("runs/demo"), max_steps=1)
            finally:
                os.chdir(old_cwd)
            qa_workdir = qa.call_args.args[0]
            self.assertTrue(qa_workdir.is_absolute())
            self.assertEqual(qa_workdir, (parent / "runs/demo").resolve())

    def test_default_step_limit_is_40(self):
        args = ppt_agent.parse_args(["--out", "runs/demo", "Build"])
        self.assertEqual(args.max_steps, 40)


@unittest.skipUnless(
    all(shutil.which(name) for name in ["soffice", "pdfinfo", "pdftotext", "pdftoppm", "montage", "unzip"]),
    "presentation QA system tools are not installed",
)
class QaIntegrationTests(unittest.TestCase):
    def _make_deck(self, root: Path, text: str = "Hello") -> None:
        from pptx import Presentation

        deck = Presentation()
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        box = slide.shapes.add_textbox(1_000_000, 1_000_000, 5_000_000, 1_000_000)
        box.text = text
        deck.save(root / "deck.pptx")

    def test_valid_deck_creates_all_artifacts(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._make_deck(root)
            old_cwd = Path.cwd()
            try:
                os.chdir(root.parent)
                passed, report = ppt_agent.qa_presentation(Path(root.name))
            finally:
                os.chdir(old_cwd)
            self.assertTrue(passed, report)
            self.assertNotIn("generator", report)
            self.assertTrue((root / "deck.pdf").is_file())
            self.assertTrue((root / "preview.jpg").is_file())
            self.assertEqual(len(list((root / "slides").glob("*.jpg"))), 1)

    def test_placeholder_fails_qa(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self._make_deck(root, "TODO replace me")
            (root / "deck.pdf").write_bytes(b"old pdf")
            (root / "preview.jpg").write_bytes(b"old preview")
            (root / "slides").mkdir()
            (root / "slides/old.jpg").write_bytes(b"old slide")
            passed, report = ppt_agent.qa_presentation(root)
            self.assertFalse(passed)
            self.assertIn("[FAIL] placeholders", report)
            self.assertEqual((root / "deck.pdf").read_bytes(), b"old pdf")
            self.assertEqual((root / "preview.jpg").read_bytes(), b"old preview")
            self.assertEqual((root / "slides/old.jpg").read_bytes(), b"old slide")

    def test_broken_pptx_fails_qa(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "deck.pptx").write_text("not a zip")
            passed, report = ppt_agent.qa_presentation(root)
            self.assertFalse(passed)
            self.assertIn("[FAIL] pptx_zip", report)


if __name__ == "__main__":
    unittest.main()
