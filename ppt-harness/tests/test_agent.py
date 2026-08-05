from __future__ import annotations

import io
import itertools
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
sys.path.append(str(ROOT / "pptx" / "scripts"))

import layout_check  # noqa: E402
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


class ContentBalanceTests(unittest.TestCase):
    def test_balanced_rectangles_pass_metric(self):
        metrics = layout_check.content_balance_metrics(
            [(0.6, 1.0, 5.5, 6.5), (7.8, 1.0, 12.7, 6.5)], 13.333, 7.5
        )
        self.assertGreater(metrics["left"], 25)
        self.assertGreater(metrics["right"], 25)
        self.assertLess(metrics["empty_vertical"], 0.25)

    def test_left_heavy_rectangles_expose_empty_band(self):
        metrics = layout_check.content_balance_metrics([(0.6, 1.0, 5.8, 6.5)], 13.333, 7.5)
        self.assertGreater(metrics["left"], 25)
        self.assertLess(metrics["right"], 5)
        self.assertGreaterEqual(metrics["empty_vertical"], 0.25)

    def test_tight_text_bounds_detect_severe_one_sided_slide(self):
        from pptx import Presentation
        from pptx.util import Inches, Pt

        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            deck = Presentation()
            deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            for row in range(9):
                box = slide.shapes.add_textbox(
                    Inches(0.65), Inches(1.35 + row * 0.5), Inches(11.4), Inches(0.35)
                )
                box.text = f"{row + 1}. Verify facts against a trusted source."
                box.text_frame.paragraphs[0].runs[0].font.size = Pt(15)
            deck.save(root / "deck.pptx")

            passed, detail = layout_check.analyse_balance(root / "deck.pptx")

            self.assertFalse(passed)
            self.assertIn("FAIL  slide 1: right body", detail)


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


class LayoutCheckTests(unittest.TestCase):
    """pptx/scripts/layout_check.py — geometry and font-metric layout checks."""

    @staticmethod
    def _deck(root: Path, build) -> Path:
        from pptx import Presentation
        from pptx.util import Inches

        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(10), Inches(5.625)
        build(deck.slides.add_slide(deck.slide_layouts[6]))
        path = root / "deck.pptx"
        deck.save(path)
        return path

    @staticmethod
    def _shape(slide, name, x, y, w, h):
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
        )
        shape.name = name
        return shape

    @staticmethod
    def _text(slide, name, x, y, w, h, text, size=12, font=None, wrap=False):
        from pptx.util import Inches, Pt

        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        box.name = name
        if wrap:
            box.text_frame.word_wrap = True
            box.text_frame.auto_size = None
        box.text_frame.text = text
        for run in box.text_frame.paragraphs[0].runs:
            run.font.size = Pt(size)
            if font:
                run.font.name = font
        return box

    def _report(self, build, waivers=None) -> str:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pptx = self._deck(root, build)
            waivers_path = None
            if waivers is not None:
                waivers_path = root / "deck.layout.json"
                waivers_path.write_text(json.dumps(waivers), encoding="utf-8")
            return layout_check.analyse(str(pptx), waivers_path)

    # -- tree building ----------------------------------------------------

    def test_background_prefix_is_never_reported(self):
        def build(slide):
            self._shape(slide, "@bg/wash", 0, 0, 10, 5.625)
            self._shape(slide, "card", 1, 1, 3, 2)

        self.assertIn("0 warnings", self._report(build))

    def test_declared_child_on_its_parent_is_not_an_overlap(self):
        def build(slide):
            self._shape(slide, "card", 0.5, 0.5, 4, 2)
            self._text(slide, "card/title", 0.7, 0.7, 3, 0.4, "Title")

        self.assertIn("0 warnings", self._report(build))

    def test_identical_rectangles_nest_by_z_order(self):
        # A number drawn on top of its circle shares the circle's rectangle.
        def build(slide):
            self._shape(slide, "Shape 0", 1, 1, 0.44, 0.44)
            self._text(slide, "Text 1", 1, 1, 0.44, 0.44, "1")

        report = self._report(build)
        self.assertIn("0 warnings", report)
        self.assertIn("named 0 | inferred 2", report)

    def test_unnamed_element_is_adopted_by_smallest_named_container(self):
        def build(slide):
            self._shape(slide, "outer", 0, 0, 9, 5)
            self._shape(slide, "card", 1, 1, 4, 2)
            self._text(slide, "Text 7", 1.2, 1.2, 1, 0.3, "hi")

        self.assertIn("0 warnings", self._report(build))

    # -- overlap ----------------------------------------------------------

    def test_partial_overlap_of_siblings_is_reported(self):
        # The slide-3 defect: dots run to 6.25, the pill starts at 6.15.
        def build(slide):
            self._shape(slide, "row/dot", 6.11, 3.24, 0.14, 0.14)
            self._shape(slide, "row/pill", 6.15, 3.16, 3.08, 0.42)

        report = self._report(build)
        self.assertIn("row/dot × row/pill", report)
        self.assertIn("71% covered", report)
        self.assertIn("[not waivable]", report)

    def test_abutting_edges_are_not_an_overlap(self):
        def build(slide):
            self._shape(slide, "a", 1, 1, 2, 1.0)
            self._shape(slide, "b", 1, 1.98, 2, 1.0)   # 0.02in sliver

        self.assertIn("0 warnings", self._report(build))

    def test_one_visual_object_is_reported_once(self):
        # A pill and the text drawn on it must not both be reported against a bar.
        def build(slide):
            self._shape(slide, "bar", 1.0, 1.0, 0.5, 2.0)
            self._shape(slide, "callout", 1.2, 1.0, 2.0, 0.5)
            self._text(slide, "callout/label", 1.2, 1.0, 2.0, 0.5, "note")

        self.assertEqual(self._report(build).count("WARN  overlap"), 1)

    # -- escape, fit, center, baseline ------------------------------------

    def test_child_escaping_its_parent_is_reported(self):
        def build(slide):
            self._shape(slide, "card", 0.5, 0.5, 2, 1)
            self._text(slide, "card/wide", 0.7, 0.7, 3, 0.4, "too wide")

        self.assertIn("sticks", self._report(build))
        self.assertIn("outside card", self._report(build))

    def test_text_too_wide_for_its_box_is_reported(self):
        # The page-number defect: 0.5in box minus 0.2in of insets holds 0.30in.
        def build(slide):
            self._text(slide, "footer/pageno", 8.95, 5.3, 0.5, 0.22,
                       "02 / 10", size=9, font="Microsoft YaHei", wrap=True)

        report = self._report(build)
        self.assertIn("footer/pageno", report)
        self.assertIn("needs 2 lines, box holds 1", report)

    def test_text_that_fits_is_not_reported(self):
        def build(slide):
            self._text(slide, "footer/pageno", 8.0, 5.3, 1.5, 0.3,
                       "02 / 10", size=9, font="Microsoft YaHei", wrap=True)

        self.assertIn("0 warnings", self._report(build))

    def test_centred_text_off_its_shape_centre_is_reported(self):
        # The badge defect: circle at y=0.64, its label at y=0.62.
        from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

        def build(slide):
            self._shape(slide, "badge", 8.85, 0.64, 0.36, 0.36)
            box = self._text(slide, "badge/mark", 8.85, 0.62, 0.36, 0.36, "?")
            box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        report = self._report(build)
        self.assertIn("badge/mark", report)
        self.assertIn("off the centre of badge", report)

    def test_row_with_mismatched_text_anchors_is_reported(self):
        # The slide-8 defect: a top-anchored title beside a middle-anchored body.
        from pptx.enum.text import MSO_ANCHOR

        def build(slide):
            self._shape(slide, "row", 0.55, 1.5, 8.9, 0.58)
            self._text(slide, "row/title", 1.42, 1.53, 1.8, 0.32, "Title", size=14)
            body = self._text(slide, "row/body", 3.3, 1.53, 5.9, 0.5, "Body", size=12)
            body.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        report = self._report(build)
        self.assertIn("WARN  baseline", report)
        self.assertIn("share a row", report)

    def test_flow_arrow_with_detached_source_is_reported(self):
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        def build(slide):
            self._shape(slide, "flow/source", 4, 0.5, 2, 1)
            self._shape(slide, "flow/target", 1, 2.5, 2, 1)
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.DOWN_ARROW, Inches(1.85), Inches(2.0), Inches(0.3), Inches(0.5)
            )
            arrow.name = "flow/arrow"

        report = self._report(build)
        self.assertIn("WARN  connector", report)
        self.assertIn("detached source endpoint", report)

    def test_flow_arrow_attached_at_both_ends_is_not_reported(self):
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches

        def build(slide):
            self._shape(slide, "flow/source", 1, 0.5, 2, 1)
            self._shape(slide, "flow/target", 1, 2.0, 2, 1)
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.DOWN_ARROW, Inches(1.85), Inches(1.5), Inches(0.3), Inches(0.5)
            )
            arrow.name = "flow/arrow"

        self.assertNotIn("WARN  connector", self._report(build))

    # -- waivers ----------------------------------------------------------

    def _overlapping_pair(self, slide):
        self._shape(slide, "chart/bar", 5.0, 0.5, 1.0, 2.0)
        self._shape(slide, "chart/callout", 5.5, 0.5, 2.0, 0.5)

    def test_waiver_silences_a_matching_overlap(self):
        report = self._report(self._overlapping_pair, waivers={"waivers": [
            {"slide": 1, "a": "chart/bar", "b": "chart/callout",
             "ratio": 0.25, "reason": "callout marks the bar"},
        ]})
        self.assertNotIn("WARN  overlap", report)
        self.assertIn("1 waived", report)

    def test_waiver_order_of_names_does_not_matter(self):
        report = self._report(self._overlapping_pair, waivers={"waivers": [
            {"slide": 1, "a": "chart/callout", "b": "chart/bar",
             "ratio": 0.25, "reason": "callout marks the bar"},
        ]})
        self.assertIn("1 waived", report)

    def test_waiver_expires_when_the_overlap_grows(self):
        report = self._report(self._overlapping_pair, waivers={"waivers": [
            {"slide": 1, "a": "chart/bar", "b": "chart/callout",
             "ratio": 0.05, "reason": "callout marks the bar"},
        ]})
        self.assertIn("waiver expired: confirmed at 5%, now 25%", report)

    def test_waiver_without_a_reason_is_refused(self):
        report = self._report(self._overlapping_pair, waivers={"waivers": [
            {"slide": 1, "a": "chart/bar", "b": "chart/callout", "ratio": 0.25},
        ]})
        self.assertIn("waiver has no reason", report)

    def test_severe_overlap_cannot_be_waived(self):
        def build(slide):
            self._shape(slide, "dot", 0.6, 3.0, 0.4, 0.4)
            self._shape(slide, "pill", 0.65, 2.9, 3.0, 0.6)

        report = self._report(build, waivers={"waivers": [
            {"slide": 1, "a": "dot", "b": "pill", "ratio": 0.9, "reason": "deliberate"},
        ]})
        self.assertIn("cannot be waived", report)
        self.assertIn("[not waivable]", report)

    def test_waiver_requires_both_elements_to_be_named(self):
        def build(slide):
            self._shape(slide, "Shape 0", 1.0, 1.0, 1.0, 2.0)
            self._shape(slide, "Shape 1", 1.5, 1.0, 2.0, 0.5)

        report = self._report(build, waivers={"waivers": [
            {"slide": 1, "a": "Shape 0", "b": "Shape 1", "ratio": 0.25, "reason": "x"},
        ]})
        self.assertIn("waivers require both elements to be named", report)

    def test_malformed_waiver_file_does_not_stop_the_run(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pptx = self._deck(root, self._overlapping_pair)
            waivers = root / "deck.layout.json"
            waivers.write_text("not json", encoding="utf-8")
            with patch("sys.stderr", new=io.StringIO()) as err:
                report = layout_check.analyse(str(pptx), waivers)
            self.assertIn("WARN  overlap", report)
            self.assertIn("deck.layout.json", err.getvalue())

    # -- reporting --------------------------------------------------------

    def test_naming_coverage_is_reported(self):
        def build(slide):
            self._shape(slide, "card", 1, 1, 2, 1)
            self._shape(slide, "Shape 9", 5, 1, 2, 1)

        report = self._report(build)
        self.assertIn("named 1 | inferred 1", report)
        self.assertIn("naming coverage 50% (1/2)", report)

    def test_missing_font_metrics_are_announced(self):
        def build(slide):
            self._text(slide, "footer/pageno", 8.95, 5.3, 0.5, 0.22,
                       "02 / 10", size=9, font="No Such Font Family", wrap=True)

        report = self._report(build)
        self.assertIn("font metrics unavailable for No Such Font Family", report)
        self.assertIn("needs 2 lines", report)   # conservative estimate still catches it

    def test_exit_code_is_zero_even_with_warnings(self):
        with tempfile.TemporaryDirectory() as temp:
            pptx = self._deck(Path(temp), self._overlapping_pair)
            with patch("sys.stdout", new=io.StringIO()) as out:
                code = layout_check.main([str(pptx)])
            self.assertEqual(code, 0)
            self.assertIn("WARN  overlap", out.getvalue())


class SkillDocumentTests(unittest.TestCase):
    """The layout check is documented in the bundled skill, not only in code."""

    def setUp(self):
        self.skill = (ROOT / "pptx" / "SKILL.md").read_text(encoding="utf-8")

    def test_script_table_stays_one_contiguous_table(self):
        # The checker's row must not be separated from the table by a blank or
        # non-row line, which would end the table in markdown.
        lines = self.skill.splitlines()
        start = next(i for i, l in enumerate(lines) if l.startswith("| Script"))
        rows = list(itertools.takewhile(lambda l: l.startswith("|"), lines[start:]))
        self.assertTrue(any("scripts/layout_check.py" in r for r in rows))

    def test_skill_documents_the_checker_and_the_naming_convention(self):
        self.assertIn("scripts/layout_check.py", self.skill)
        self.assertIn("objectName", self.skill)
        self.assertIn("@bg/", self.skill)
        self.assertIn("deck.layout.json", self.skill)


if __name__ == "__main__":
    unittest.main()
