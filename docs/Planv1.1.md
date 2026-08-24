Implement PowerPoint MCP v1.1.

This is the next major feature iteration after the successful v1 priority fixes.

The v1 foundation is now considered stable:

- active-session / working-copy safety is fixed
- save/save_as behavior is reliable
- high-fidelity PowerPoint COM rendering works
- targeted text inspection exists
- batch text mutation exists
- batch shape mutation exists
- paragraph/bullet preservation has been improved
- validation has basic container-awareness
- inspection output is substantially more agent-friendly

The next objective is NOT simply adding more low-level tools.

v1.1 should introduce higher-level PowerPoint semantics so an agent can reason about:
- typography hierarchy
- cards/containers
- groups of related objects
- relative layout changes
- semantic styling
- actual text fit
- repeated visual structures
- slide-wide design consistency

The guiding principle for v1.1 is:

"Move complexity from the agent into deterministic PowerPoint-specific operations, while keeping the agent in control of visual and semantic decisions."

Do not introduce an LLM into the MCP.
Do not make the MCP itself autonomous.
Do not replace the existing rendering/verification loop.

==================================================
V1.1 ARCHITECTURE
==================================================

Organize the v1.1 capabilities into five layers:

1. INSPECTION
2. SEMANTIC STRUCTURE
3. TYPOGRAPHY / STYLING
4. LAYOUT OPERATIONS
5. VERIFICATION

The agent should be able to move from:

inspect raw shapes
→ understand visual structure
→ perform high-level deterministic transformation
→ render
→ validate
→ visually correct
→ save

without having to write ad-hoc Python scripts for common PowerPoint operations.

==================================================
FEATURE 1 — RELATIVE TYPOGRAPHY OPERATIONS
==================================================

Extend ppt_modify_text and ppt_batch_modify_text to support relative font changes.

Currently the agent must calculate absolute sizes manually.

Add support for:

font_size_delta

and:

font_size_scale

Examples:

{
  "shape_id": 17,
  "font_size_delta": 2
}

or:

{
  "shape_id": 17,
  "font_size_scale": 1.15
}

Rules:

- absolute font_size takes precedence if explicitly supplied
- delta applies to the existing size
- scale applies to the existing size
- do not allow negative or invalid resulting sizes
- support min/max bounds where useful

For batch operations:

{
  "shape_id": 17,
  "font_size_delta": 2,
  "min_font_size": 12,
  "max_font_size": 24
}

Return the original and resulting font sizes.

Do not break existing absolute font-size behavior.

Add tests for:
- absolute size
- delta
- scale
- min/max
- mixed batch operations
- multiple runs
- bullet text

==================================================
FEATURE 2 — ppt_scale_slide_typography
==================================================

Implement a high-level deterministic typography operation:

ppt_scale_slide_typography

Purpose:

Allow the agent to proportionally increase/decrease typography on an entire slide while preserving hierarchy.

Example:

ppt_scale_slide_typography(
    slide_number=2,
    scale_factor=1.15,
    min_pt=9,
    max_pt=30
)

IMPORTANT:

This is NOT simply:

font_size = font_size * scale_factor

for every text run.

The implementation must:

- inspect all text-bearing shapes
- preserve relative hierarchy
- respect min/max sizes
- preserve paragraph formatting
- preserve bullet formatting
- avoid modifying intentionally tiny UI elements where possible
- return a summary of affected shapes

Where semantic roles are available, use them.

If role information is unavailable, use conservative heuristics based on:
- shape geometry
- existing font size
- boldness
- position
- text length
- containment
- repetition

Do not claim that this operation guarantees perfect layout.

After scaling, the agent should be expected to validate and render the slide.

The tool should report:
- shapes modified
- old/new sizes
- shapes skipped
- reasons for skipping

==================================================
FEATURE 3 — SEMANTIC SHAPE / TEXT ROLE DETECTION
==================================================

Introduce a lightweight semantic-role layer.

The MCP currently exposes shape IDs and geometry, but the agent repeatedly has to infer:

- title
- subtitle
- section header
- card title
- body
- bullet
- badge
- metric
- footer
- background
- icon
- connector

Add a deterministic role inference mechanism.

Do NOT use an LLM.

Implement:

ppt_analyze_slide_structure

The result should identify likely roles and confidence.

Example:

{
  "slide_number": 2,
  "elements": [
    {
      "shape_id": 12,
      "role": "slide_title",
      "confidence": 0.96
    },
    {
      "shape_id": 27,
      "role": "card",
      "confidence": 0.91
    },
    {
      "shape_id": 28,
      "role": "card_title",
      "parent": 27,
      "confidence": 0.88
    }
  ]
}

Do not over-engineer this.

The first version should focus on useful deterministic roles:

- background
- slide_title
- subtitle
- section_header
- card
- card_title
- body
- bullet
- badge
- metric
- footer
- connector
- icon

Roles should be treated as inferred metadata, not absolute truth.

The agent must be able to override or ignore them.

==================================================
FEATURE 4 — CONTAINER / CARD HIERARCHY
==================================================

The latest real-world review shows that the MCP needs to understand structures such as:

Card
 ├── icon
 ├── title
 ├── bullets
 └── metric

The agent currently has to manually associate shape IDs.

Introduce a structural abstraction for containment.

Implement:

ppt_analyze_containers

or incorporate container analysis into:

ppt_analyze_slide_structure

The system should identify likely parent-child relationships using:

- geometric containment
- z-order
- fill/background characteristics
- proximity
- repeated layout patterns
- shape types
- text position

Example:

Card 41
  ├── Icon 42
  ├── Title 43
  ├── Body 44
  └── Metric 45

Do not physically group the PowerPoint objects unless explicitly requested.

This is a logical MCP-level hierarchy.

Return confidence scores.

The hierarchy must be usable by later operations.

==================================================
FEATURE 5 — CONTAINER-AWARE LAYOUT OPERATIONS
==================================================

Add high-level operations that operate on logical containers.

At minimum:

ppt_resize_container
ppt_move_container
ppt_reflow_container

The logical container may consist of:
- background
- title
- body
- icon
- metric
- decorative elements

Example:

ppt_resize_container(
    container_id=41,
    width=4.2,
    height=3.8
)

The MCP should:

1. identify children
2. resize/reposition the parent
3. preserve relative child layout where possible
4. keep padding consistent
5. prevent obvious child overflow
6. return resulting geometry

This should be deterministic.

Do not attempt arbitrary responsive-layout intelligence.

Focus on common card-based presentation layouts.

==================================================
FEATURE 6 — LAYOUT PRIMITIVES
==================================================

Introduce high-level deterministic layout operations.

At minimum:

ppt_align_shapes
ppt_distribute_shapes
ppt_space_shapes
ppt_equalize_sizes

Support operations such as:

- align_left
- align_center
- align_right
- align_top
- align_middle
- align_bottom
- distribute_horizontal
- distribute_vertical
- equal_width
- equal_height
- equal_size

These should operate on shape IDs.

Example:

{
  "slide_number": 3,
  "shape_ids": [13,14,15,16],
  "operation": "distribute_vertical"
}

The goal is to eliminate agent-side arithmetic.

The MCP should calculate exact coordinates.

These operations must preserve the active session model.

==================================================
FEATURE 7 — THEME / ROLE-BASED STYLE PRESETS
==================================================

Implement a lightweight role-based styling mechanism.

Do NOT build a complete PowerPoint theme engine.

Support deterministic style presets such as:

slide_title
section_header
card_title
body
bullet
badge
metric
footer

Implement:

ppt_apply_style

Example:

{
  "shape_ids": [14,18,21],
  "style_role": "card_title"
}

Or:

{
  "role": "card_title",
  "font_size": 14,
  "bold": true
}

Where possible, derive style values from existing deck styles rather than hard-coding arbitrary corporate design choices.

Allow a reference shape:

style_source_shape_id

This is particularly important for preserving an existing deck's visual language.

Example:

ppt_apply_style(
    shape_ids=[41,42,43],
    style_source_shape_id=18
)

The new styling should copy supported properties such as:

- font
- font size
- weight
- color
- paragraph alignment
- margins
- fill
- line
- line width

Do not modify geometry unless explicitly requested.

==================================================
FEATURE 8 — STYLE INHERITANCE FOR NEW OBJECTS
==================================================

Extend the shape creation capabilities from the previous iteration.

ppt_add_shape
ppt_add_textbox
ppt_add_connector

must support:

style_source_shape_id

Example:

Create a new card using the existing card as the style source.

The new object should inherit relevant visual properties while getting its own geometry/text.

This is critical for editing an existing presentation because the agent should not need to manually reconstruct:
- colors
- borders
- fonts
- line widths
- transparency
- paragraph settings

from scratch.

==================================================
FEATURE 9 — REAL TEXT FIT / OVERFLOW MEASUREMENT
==================================================

This is now a high-priority validation improvement.

Current heuristic:

character-count based overflow estimation

is producing false positives such as:

"59% text overflow"

when the text actually fits comfortably on a wide banner.

Do not rely on character count as the primary overflow detector.

Because PowerPoint COM rendering is already available, implement a more accurate measurement strategy on Windows.

Use actual PowerPoint text-frame information where possible, such as:

- TextFrame/TextFrame2 dimensions
- AutoSize behavior
- text range dimensions
- font metrics available through PowerPoint
- actual rendered bounds where available

The goal is to determine whether text genuinely exceeds its text box.

The validator should distinguish:

- fits
- likely overflow
- confirmed overflow
- unable to determine

Do not report a precise percentage such as "59% overflow" unless that percentage is actually derived from a reliable measurement.

Keep a fallback heuristic for environments without COM, but clearly identify it as heuristic.

Add tests covering:
- single-line wide text
- multiline text
- long bullet
- narrow card
- large title
- intentional clipping
- text with different fonts

==================================================
FEATURE 10 — SMART TINY-FONT DETECTION
==================================================

Current validator flags intentionally compact badge/pill labels as suspiciously tiny.

Do not simply remove the tiny-font rule.

Make it context-aware.

Distinguish:

Potentially problematic:
- body text at 7pt
- bullet text at 7pt
- major content at 8pt

from:

Potentially intentional:
- badge
- pill
- small metadata
- footer
- compact label

Use semantic role, geometry, position, and surrounding structure.

Return:

CRITICAL_TINY_TEXT
SUSPICIOUS_TINY_TEXT
INTENTIONAL_COMPACT_TEXT

where appropriate.

The validator should remain conservative.

The goal is to reduce noise without hiding actual readability problems.

==================================================
FEATURE 11 — IMPROVE VAL-01 CONTAINMENT
==================================================

The previous improvement was not sufficient.

The latest test still produced:

58 overlap warnings

where text boxes inside rounded card backgrounds were classified as ACTUAL_OVERLAP.

Fix the hierarchy model so the validator can understand:

container
  └── content

as intentional.

Use:

- z-order
- geometric containment
- fill
- shape type
- proximity
- inferred parent-child relationships

Do not suppress all overlaps.

The validator must continue detecting:

- independent card collisions
- overlapping nodes
- text crossing unrelated objects
- connector collisions
- objects extending outside containers

The goal is for VAL-01 to become a trustworthy signal for the agent.

==================================================
FEATURE 12 — COMPOSITE FLOW DIAGRAM TOOL
==================================================

Now implement the previously deferred high-level diagram operation:

ppt_create_flow_diagram

This should create a deterministic native PowerPoint flow diagram from structured input.

Example:

{
  "slide_number": 5,
  "steps": [
    "Request",
    "Validate",
    "Analyze",
    "Plan",
    "Approve",
    "Configure",
    "Deploy",
    "Test",
    "Monitor",
    "Complete"
  ],
  "direction": "horizontal",
  "columns": 5,
  "rows": 2
}

Support at minimum:

- steps
- rows
- columns
- direction
- node size
- spacing
- connector type
- style source

The MCP should:

1. calculate geometry
2. create nodes
3. create connectors
4. align nodes
5. distribute nodes
6. apply style
7. return all created IDs
8. return the resulting bounding box

The operation must use native PowerPoint shapes and connectors.

Do not create a flattened image.

The result must remain editable.

This is specifically intended to turn:

"Create a 10-step flow diagram"

from potentially 20+ low-level MCP calls into one deterministic composite operation.

==================================================
FEATURE 13 — COMPOSITE SLIDE EDIT OPERATIONS
==================================================

Do not create a giant universal "edit slide" tool.

Instead, identify a small number of useful composite operations.

Candidates:

ppt_create_flow_diagram
ppt_resize_container
ppt_reflow_container
ppt_apply_style
ppt_scale_slide_typography

Keep primitive tools available.

The design principle is:

- primitive tools for unusual operations
- composite tools for common repeated presentation tasks

Do not expose dozens of highly specialized macros.

==================================================
FEATURE 14 — VISUAL VERIFICATION METADATA
==================================================

Improve ppt_render_slide response metadata.

The render itself remains the source of truth for visual appearance.

The MCP response should additionally report:

- rendered dimensions
- slide number
- render path
- render timestamp
- whether the slide changed since the previous render
- optionally a render hash

Avoid rerendering an unchanged slide when practical.

Do not replace visual inspection with automated heuristics.

The intended loop remains:

modify
→ render
→ inspect image
→ modify if necessary

==================================================
FEATURE 15 — STRUCTURAL DIFF
==================================================

Improve ppt_compare_slides so that the agent can understand what changed.

Return concise structural information:

- added shapes
- deleted shapes
- moved shapes
- resized shapes
- text changes
- typography changes
- style changes

Example:

{
  "added": [101,102],
  "deleted": [27,28],
  "moved": [13,14],
  "resized": [13,18],
  "text_changed": [34,36,38],
  "typography_changed": [34,36]
}

Keep detailed information available separately.

==================================================
FEATURE 16 — AGENT SKILL UPDATE
==================================================

Update SKILL.md substantially.

The skill should teach this decision hierarchy:

FIRST:
Use targeted inspection.

SECOND:
Use semantic analysis when the task involves layout structure.

THIRD:
Prefer batch operations for repeated changes.

FOURTH:
Prefer high-level deterministic operations for common patterns.

FIFTH:
Render the affected slide.

SIXTH:
Validate.

SEVENTH:
Make corrective edits if necessary.

EIGHTH:
Save.

Examples:

Task:
"Make all the text bigger."

Prefer:
ppt_inspect_text
→ reason about hierarchy
→ ppt_scale_slide_typography
→ validate
→ render

Task:
"Make these four cards the same width."

Prefer:
ppt_equalize_sizes
or appropriate layout primitive.

Task:
"Move the entire card down."

Prefer:
ppt_move_container

Task:
"Create a 10-step process."

Prefer:
ppt_create_flow_diagram

Task:
"Change these 20 bullet texts."

Prefer:
ppt_batch_modify_text

Task:
"Match the style of this existing box."

Prefer:
style_source_shape_id / ppt_apply_style

Explicitly prohibit unnecessary external Python scripts when an MCP operation already supports the required action.

External scripts remain acceptable only when:
- the MCP genuinely lacks the capability
- the task requires unsupported low-level PowerPoint behavior
- the agent is diagnosing an MCP implementation problem

==================================================
V1.1 TEST STRATEGY
==================================================

Create a dedicated v1.1 test suite.

Test categories:

1. Typography
2. Containers
3. Layout
4. Styling
5. Diagram creation
6. Validation
7. Rendering
8. Session safety
9. Structural diff

Test real PPTX files, not only mocked objects.

==================================================
REAL-WORLD BENCHMARKS
==================================================

Re-run the following real-world scenarios.

SCENARIO 1:

"Update slide 2. Clean up the text, and make it appropriate size for the slide. It is too small everywhere on this slide."

Expected approximate workflow:

inspect_text
→ analyze structure
→ typography operation / batch text mutation
→ validate
→ render
→ correction if required

SCENARIO 2:

"On Slide 3. remove the Orchestration box, and its adjacent box. Extend the Client Configuration box to the bottom of the slide, and also add content and extend the box adjacent to the Client Configuration box."

Expected workflow should use:
- structural analysis
- container/layout operations
- batch mutations
- render
- validation

SCENARIO 3:

"Create a slide with a 10-step flow diagram using the provided steps."

Expected workflow should use:

ppt_create_flow_diagram
→ render
→ validate
→ save

rather than dozens of low-level calls.

SCENARIO 4:

"Make these five cards match the style of this existing card and distribute them evenly."

Expected workflow should use:

structure analysis
→ style inheritance
→ equalize/distribute
→ render
→ validate

==================================================
SUCCESS METRICS
==================================================

Measure:

1. MCP tool calls
2. mutation calls
3. inspection calls
4. render calls
5. validation calls
6. agent-generated external scripts
7. response size
8. corrective iterations
9. final visual correctness

The key metric is not minimum tool calls.

The key metric is:

"How much useful PowerPoint work can one MCP call safely accomplish?"

v1.1 should significantly reduce:
- repetitive geometry arithmetic
- repetitive font arithmetic
- manual shape-ID association
- external filtering scripts
- low-level object creation
- false validation noise

while preserving:
- deterministic behavior
- native editable PowerPoint objects
- session isolation
- automatic backups
- high-fidelity rendering
- visual verification

==================================================
IMPORTANT DESIGN CONSTRAINTS
==================================================

1. Do not put an LLM inside the MCP.

2. Do not make the MCP decide what the presentation "should" look like.

3. The agent remains responsible for semantic decisions.

4. The MCP is responsible for precise deterministic execution.

5. Do not flatten slides into images.

6. Preserve native PowerPoint objects.

7. Preserve existing formatting unless explicitly changed.

8. Preserve bullet and paragraph formatting.

9. Do not silently modify the source presentation outside the active session.

10. Every mutation must operate on the session working copy.

11. Every composite operation must be reversible through the existing session/revert system.

12. Maintain the existing backup behavior.

13. Do not create a giant universal PowerPoint editing API.

14. Prefer a small number of high-value composable operations.

15. All high-level operations must still be inspectable and verifiable through the existing rendering loop.

==================================================
IMPLEMENTATION ORDER
==================================================

Implement in this order:

PHASE A — TYPOGRAPHY
1. relative font operations
2. ppt_scale_slide_typography
3. smart tiny-font validation
4. real text-fit measurement

PHASE B — STRUCTURE
5. ppt_analyze_slide_structure
6. container hierarchy
7. container-aware validation

PHASE C — LAYOUT
8. alignment
9. distribution
10. equal sizing
11. container move/resize/reflow

PHASE D — STYLE
12. style inheritance
13. ppt_apply_style
14. role-based style presets

PHASE E — COMPOSITE CREATION
15. ppt_create_flow_diagram

PHASE F — VERIFICATION
16. improved structural diff
17. render metadata / caching
18. end-to-end benchmarks

After each phase:
- run unit tests
- run relevant real-PPTX tests
- do not proceed if existing functionality regresses

==================================================
FINAL DELIVERABLE
==================================================

Provide:

1. Complete list of new tools.
2. Modified tools.
3. Updated SKILL.md.
4. Internal architecture changes.
5. Tests added.
6. Real PPTX test results.
7. Before/after tool-call benchmarks.
8. Before/after response sizes.
9. Number of external scripts required before vs after.
10. Screenshots/renders for representative scenarios.
11. Known limitations.
12. Recommended v1.2 features.

Do not declare v1.1 complete merely because the unit tests pass.

The system must be exercised against actual PowerPoint presentations using the full:

INSPECT
→ ANALYZE
→ MUTATE
→ VALIDATE
→ RENDER
→ CORRECT
→ SAVE

workflow.