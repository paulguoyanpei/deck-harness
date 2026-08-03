---
name: pptx-api
description: Generate a complete, self-contained pptxgenjs presentation source file from a presentation brief in one model response. Use when the model has no tools, filesystem, execution environment, rendering, or iterative QA capability and must return source code that a caller can save and run later.
---

# One-shot PPTX source generation

Generate a polished presentation as one complete JavaScript source file. Assume no tools are available: do not read files, run commands, fetch URLs, render slides, inspect images, or revise an artifact after execution.

## Output contract

- Return exactly one fenced `javascript` code block and no prose outside it.
- Make the code self-contained and directly runnable after the caller installs `pptxgenjs`.
- Write the presentation to `deck.pptx` with `await pptx.writeFile({ fileName: "deck.pptx" })`.
- Do not use local files, bundled resources, shell commands, generated images, remote images, or packages other than `pptxgenjs`.
- Use native PowerPoint text, shapes, lines, tables, and charts. Build simple icons and diagrams from shapes.
- Do not claim that the code or deck was executed, rendered, opened, inspected, or validated.
- Do not emit placeholders, TODOs, omitted sections, pseudocode, or comments asking the caller to finish content.
- Include all requested slides and all substantive copy in the returned source.

## Plan internally before writing

Infer the audience, purpose, tone, slide count, narrative arc, and appropriate level of detail from the brief. Create a coherent story:

1. Establish the promise or central question.
2. Develop the argument through distinct, logically ordered sections.
3. Use evidence, comparisons, processes, timelines, or frameworks where they clarify the story.
4. End with a synthesis, recommendation, or memorable takeaway.

Give each slide one job and one clear takeaway. Prefer concise, specific writing over generic business language. Do not repeat the same facts across slides.

## Visual direction

- Choose a content-specific palette. Give one color 60–70% of the visual weight, use one or two supporting colors, and reserve a sharp accent for emphasis.
- Establish one recurring visual motif, such as rounded image-like panels, numbered circles, or a distinctive diagram language.
- Vary layouts across the deck: hero, split composition, comparison, process, timeline, metric callouts, chart, and conclusion.
- Put at least one meaningful visual element on every slide using native shapes, diagrams, or charts.
- Avoid repetitive title-and-bullets layouts, decorative stripes, title underlines, generic blue defaults, and cream/beige defaults.
- Use dark/light contrast deliberately. Title and conclusion slides may be dark while content slides remain light.
- Keep a minimum 0.5-inch outer margin and 0.3–0.5-inch gaps between content blocks.
- Left-align body copy. Center only short titles, numbers, or intentional hero elements.
- Keep body copy short enough to fit. Split dense material across slides instead of shrinking it excessively.

## Typography

- Use Arial, Calibri, Cambria, Times New Roman, Courier New, Bookman Old Style, or Century Schoolbook.
- Never default to Aptos.
- Use 36–44 pt bold slide titles, 20–24 pt section headers, 14–18 pt body text, and 10–12 pt captions.
- Use no more than two font families and maintain strong size contrast.
- Set `margin: 0` when text must align exactly with shapes or other text.

## PptxGenJS correctness

- Create one `pptxgenjs` instance and set `pptx.layout = "LAYOUT_WIDE"` before adding slides.
- Use the 13.333 × 7.5-inch wide canvas and keep every object within it.
- Use six-digit hex colors without `#` or embedded alpha values.
- Apply translucency with `transparency` on fills and images and `opacity` on shadows.
- Create a fresh options object for every shape, text, image, and chart call; pptxgenjs mutates option objects.
- Use non-negative shadow offsets. Use an angle to change shadow direction.
- Use `charSpacing`, not `letterSpacing`.
- Use `pptx.ShapeType.roundRect` for rounded rectangles.
- Do not request unsupported gradient fills; approximate depth with layered solid or translucent shapes.
- Never use literal bullet characters. For true lists, use rich-text runs with `bullet` and `breakLine` options.
- Add speaker notes only when requested, using `slide.addNotes("...")` once per slide.
- Keep charts native with `slide.addChart()`. Supply titles, data labels where useful, a deliberate palette, quiet axes/gridlines, and a legend only when it adds information.
- For stacked bars or columns, use only `ctr`, `inEnd`, or `inBase` for `dataLabelPosition`.
- Avoid secondary axes unless essential. If used, declare both value and category axes completely.
- Never reorder or post-process presentation XML.

## Code organization

- Start with `const pptxgen = require("pptxgenjs");`.
- Define theme colors, typography, dimensions, and small helper functions near the top.
- Keep helpers deterministic and free of external dependencies.
- Add slide numbers discreetly unless the brief calls for a different footer treatment.
- Add a short `warnIfSlideHasOverlaps(slide, pptx)`-style static check only if it is fully implemented from the objects created in the code; otherwise omit it rather than pretending to validate.
- End with an async `main()` that writes `deck.pptx` and reports errors through `console.error` and a nonzero exit code.

## Content and visual self-check before responding

Review the source mentally before emitting it:

- Confirm the requested slide count and topic coverage.
- Confirm that every slide has a distinct title, takeaway, and visual composition.
- Confirm that coordinates and dimensions stay within the wide slide canvas.
- Confirm that repeated components align consistently and retain adequate gaps.
- Confirm that text volumes are plausible for their boxes and font sizes.
- Confirm that contrast is sufficient and no placeholder content remains.
- Confirm that all identifiers, strings, arrays, callbacks, and async calls form valid JavaScript.

This is a static reasoning check only. Never describe it as runtime, structural, or visual validation.
