# PowerPoint MCP v1.2 — Benchmark Plan

Sample file: `benchmark_deck.pptx` (20 slides, one target section per task below).
Use a **fresh, untouched copy of this file per agent** — the server keeps working
copies under `.ppt-agent/sessions/<session_id>/`, and running two agents against
the same source path at the same time risks session/backup collisions. Practically:
`cp benchmark_deck.pptx runs/task-01/benchmark_deck.pptx`, one subfolder per task,
then point each agent's `presentation_path` at its own copy.

---

## 1. One-shot prompt template

Fill in the four bracketed fields and hand the whole thing to a fresh Antigravity
agent session (one per task, run in parallel).

```
Use the powerpoint-mcp server to open [FILE_PATH] and complete the following task on slide(s) [SLIDE_NUMBER(S)]:

[TASK INSTRUCTION]

Work only within the standard tool workflow: inspect before mutating, use PATCH
semantics (don't touch fields you don't need to change), and re-render + validate
the slide(s) you touched when you're done. Do not modify any slide other than the
one(s) named above. Save the result as [OUTPUT_FILE_PATH].

When you are finished, report back in exactly this format:

## Task Report: [TASK NAME]
- Task: <one-line restatement of what you were asked to do>
- Tool calls (total): <count>
- MCP calls: <count and list of tool names called, in order, with call count each>
- Agent-generated scripts: <count and, for each, language + one-line purpose — 0 if none>
- Inspection calls: <count — e.g. ppt_inspect_slide, ppt_inspect_shape, ppt_analyze_containers, etc.>
- Mutation calls: <count — e.g. ppt_modify_text, ppt_move_shape, ppt_batch_modify_table_cells, etc.>
- Rendering calls: <count — ppt_render_slide / ppt_render_slides / ppt_render_presentation>
- Validation calls: <count — ppt_validate_slide / ppt_validate_slides, plus results (pass/fail + rule IDs triggered)>
- Retries: <count of any tool call repeated after a failure or an unsatisfactory result, and why>
- Final quality (self-assessed, 1-5): <score> — <one or two sentences justifying it: did validation pass clean, does the render look correct, was anything on an untouched slide altered>
```

---

## 2. The 15 tasks

| # | Task | Slide(s) | Objective |
|---|---|---|---|
| 1 | Change all titles on a slide | 2 | Rename the slide title and all three card titles ("Reliability", "Velocity", "Scale") to new names you choose, without touching body text. |
| 2 | Increase typography across a slide | 3 | Scale up every text element's font size proportionally (title, subtitle, two body paragraphs, caption) while preserving the existing size hierarchy between them. |
| 3 | Add an image | 4 | Insert an image into the empty area to the right of the bullet list (roughly 4.5in × 3.2in, starting around x=7.6in, y=2.0in), preserving its aspect ratio. |
| 4 | Replace an image | 5 | Replace the existing customer photo with a new image, keeping the exact position and size of the original. |
| 5 | Populate a 5×5 table | 6 | Fill in the four empty data rows under the existing header row with realistic values for each column (Region, Revenue, Growth, NPS, Status). |
| 6 | Reformat a table | 7 | Restyle the existing vendor-comparison table: bold header row with a fill color, alternating row shading, consistent alignment, and column widths that fit the content without wrapping. |
| 7 | Move a card | 8 | Move "Card_Beta" so it's vertically aligned with "Card_Alpha" and "Card_Gamma" in the same row. |
| 8 | Resize a card | 9 | Resize "Card_Two" so its width and height match "Card_One" and "Card_Three", without moving its top-left position more than necessary. |
| 9 | Harmonize four slides | 10–13 | Using slide 10 as the reference, sync the header (category label + title style), footer, and stepper across slides 11, 12, and 13 — without changing each slide's own body text. |
| 10 | Create a 10-step flow | 14 | Build a 10-step horizontal flow diagram below the title for a deployment pipeline: Commit, Build, Unit Tests, Security Scan, Package, Stage Deploy, Integration Tests, Approval, Prod Deploy, Monitor. |
| 11 | Modify an existing flow | 15 | Update the existing 5-step onboarding stepper so "Configure" (step 3) is now the active step and "Sign Up" and "Verify" are marked completed. |
| 12 | Add/remove content from a card | 16 | On the "Growth" plan card, remove the "SSO (SAML)" bullet and add two new bullets: "Priority email support" and "Usage analytics dashboard". |
| 13 | Rebuild a table-heavy slide | 17 | Rebuild this slide so the three overlapping/off-slide tables become one clean, non-overlapping layout that fits entirely within the slide bounds and reads clearly. |
| 14 | Make two slides visually consistent | 18–19 | Restyle slide 19 ("Team B Update") to visually match slide 18's typography, color palette, and card spacing, while keeping slide 19's own text content unchanged. |
| 15 | Fix a visually broken slide | 20 | Fix all validation issues on this slide: the card overlapping the title, the 6pt body text, the caption overflowing its box, the card running off the right edge of the slide, and the two overlapping colored cards. |

---

## 3. Aggregate comparison sheet

After all 15 task reports come back, roll them into one table for the v1.2
retro (and to diff against future versions run on the same tasks):

| # | Task | Total tool calls | MCP calls | Scripts | Inspection | Mutation | Render | Validate | Retries | Quality (1-5) | Notes |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | | | | | | | | | | | |
| ... | | | | | | | | | | | |

Things worth eyeballing once the sheet is full:
- **Tasks with high retry counts** — usually where a tool's parameters are
  ambiguous or a PATCH didn't behave as expected; good candidates for the next
  round of fixes.
- **Tasks where the agent wrote a custom script instead of using an MCP tool**
  — signals a missing or under-powered tool for that operation.
- **Validation calls skipped entirely** — the agent declared success without
  checking; worth flagging even if the result happened to look fine.
- **Quality scores below 4** — read the agent's own justification first; it
  usually points at the specific rule or visual defect.

---

## 4. Running this in parallel

1. Make one subfolder per task under `runs/`, each with its own copy of
   `benchmark_deck.pptx`.
2. Spin up 15 Antigravity sessions (or however many you can run concurrently),
   one per task, each with its filled-in prompt from Section 1.
3. Collect the 15 "Task Report" blocks verbatim — don't paraphrase them — and
   drop each row into the Section 3 sheet.
4. Re-run the same 15 tasks against the next MCP version on fresh copies of the
   same source deck, and diff the two aggregate sheets directly.
