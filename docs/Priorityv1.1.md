Implement the next focused iteration of the PowerPoint MCP.

This iteration is NOT a broad v1.1 implementation. The goal is to fix the highest-priority agent-ergonomics problems discovered during two real-world editing tasks.

The current MCP is already capable of:
- isolated session-based editing
- non-destructive working copies
- high-fidelity PowerPoint COM rendering
- validation
- precise geometry editing
- text editing
- save/revert
- visual verification

Do not redesign those systems.

Implement the following six improvements, in this order:

==================================================
PRIORITY 1 — CONCISE / FILTERED SLIDE INSPECTION
==================================================

Problem observed:

ppt_inspect_slide(detail="full") can return more than 8,100 lines / ~208 KB on a complex slide with ~96 shapes.

This forces the agent to consume excessive context and sometimes write an external filtering script just to find text shapes.

Goal:

Make inspection agent-friendly without removing any existing information.

Implement:

1. Keep existing full inspection capability.

2. Add or formalize inspection detail levels:

   detail="summary"
   detail="full"

   Default must be "summary".

3. Add filtering capabilities, at minimum:

   text_only=true/false

   and preferably useful filters such as:

   include_geometry=true/false
   include_style=true/false
   include_xml=true/false
   include_images=true/false

4. The default summary for a slide should contain:

   - slide number
   - slide dimensions
   - total shape count
   - meaningful shape IDs
   - semantic roles
   - x/y/width/height
   - concise text
   - relevant font information
   - major warnings

5. Do NOT include raw XML, complete run-level information, relationship data, image metadata, or other deep diagnostics in the default summary unless necessary.

6. `detail="full"` must continue exposing the existing detailed information.

7. Filtering must happen inside the MCP server.

   DO NOT require the LLM to:
   - retrieve a huge result
   - write a Python script
   - filter the result externally.

8. Make the filtering response structured and concise.

Example useful response for text_only=true:

Slide 2
Text objects: 25

[12] title
  text: "..."
  role: title
  font: Aptos Display 28pt
  bbox: x=0.72 y=0.31 w=11.8 h=0.55

[18] body
  text: "..."
  role: body
  font: Calibri 9pt
  bbox: ...

The detailed inspection API must remain available when needed.

Add tests for:
- summary mode
- full mode
- text_only
- combined filters
- complex slide output size
- preservation of existing detailed data

==================================================
PRIORITY 2 — ADD ppt_inspect_text
==================================================

Create a dedicated agent-facing inspection tool:

ppt_inspect_text

Purpose:

Efficiently inspect all textual content on a slide without forcing the agent to parse the entire shape tree.

Inputs should support at minimum:

- slide_number
- include_geometry
- include_style
- include_paragraph_metadata

Return, for each text-bearing shape:

- shape_id
- semantic role
- text
- font family
- font size
- bold/italic where relevant
- text color where relevant
- x/y/width/height
- overflow/boundary status where available
- bullet/paragraph metadata where relevant

Keep the output concise.

This tool should be ideal for prompts such as:

"Clean up the text on slide 2 and make it a reasonable size."

The agent should be able to:

ppt_inspect_text
    ↓
reason about all text
    ↓
batch modify
    ↓
validate
    ↓
render

without having to inspect all non-text shapes.

Add tests.

==================================================
PRIORITY 3 — ADD ppt_batch_modify_text
==================================================

Create:

ppt_batch_modify_text

Purpose:

Allow the agent to modify many text shapes in a single MCP call.

This is now a critical capability.

A slide may contain 20–30 text shapes that need changes.

The agent should NOT need:

ppt_modify_text × 25

Example conceptual input:

{
  "slide_number": 2,
  "operations": [
    {
      "shape_id": 12,
      "text": "Updated text",
      "font_size": 22
    },
    {
      "shape_id": 18,
      "text": "Updated body copy",
      "font_size": 16
    },
    {
      "shape_id": 22,
      "font_size": 14
    }
  ]
}

Each operation should support the safe subset of the existing ppt_modify_text properties.

Requirements:

- execute all mutations against the active session working copy
- validate shape IDs before modifying
- return concise per-operation success/failure
- preserve unrelated formatting
- do not partially corrupt the presentation
- use the same target-resolution/session safety logic already implemented
- do not bypass session isolation

Where practical, validate the entire operation list before making mutations.

If a single operation is invalid, prefer failing the batch safely instead of leaving a partially applied batch.

Do not duplicate the existing text-editing implementation unnecessarily.

Refactor shared mutation logic into reusable internal functions if appropriate.

Add comprehensive tests.

==================================================
PRIORITY 4 — FIX PARAGRAPH / BULLET FORMAT PRESERVATION
==================================================

Current problem:

Normal text replacement can unintentionally strip paragraph-level formatting such as:

- bullet characters
- bullet fonts
- left margins
- hanging indents
- negative indents
- paragraph spacing
- alignment
- other paragraph-level XML

This is unacceptable for precision PowerPoint editing.

Change the default text-editing behavior so that:

WHEN USER CHANGES TEXT ONLY:

Preserve all existing paragraph formatting unless the user explicitly requests paragraph formatting changes.

Preserve, where present:

- bullet character
- bullet font
- numbering/bullet properties
- left margin
- hanging indent
- indentation
- paragraph spacing
- line spacing
- alignment
- paragraph-level properties

Also preserve run-level formatting whenever applicable.

Do not require the LLM to explicitly pass a "preserve_bullets=true" flag for normal operations.

The default behavior must be preservation.

Both:
- ppt_modify_text
- ppt_batch_modify_text

must use the same preservation logic.

Test cases must include:

1. bullet text replacement
2. multiple bullet levels
3. indentation
4. hanging indents
5. rich text with multiple runs
6. text replacement without formatting changes
7. explicit formatting changes when requested

The goal is:

"Change the words, keep the formatting."

==================================================
PRIORITY 5 — CONTAINER-AWARE OVERLAP VALIDATION
==================================================

Current problem:

ppt_validate_slide produces a very large number of false VAL-01 overlap warnings because common PowerPoint designs intentionally place:

- text frames
- icons
- labels

on top of filled rectangular background/card shapes.

The current validator interprets these legitimate containment relationships as collisions.

Improve VAL-01.

At minimum, recognize a likely container/content relationship when:

- a background shape has a visible fill
- the content shape is above it in z-order
- the content is spatially contained within the background shape
- the content type is a text frame/icon/content element
- the background does not appear to be another independent diagram element

Do not blindly ignore all overlaps.

Distinguish:

VALID CONTAINMENT:
background card
    └── title
    └── body
    └── icon

from:

ACTUAL COLLISION:
card A overlaps card B
two independent diagram nodes overlap
arrow crosses unintended content
text box extends outside its intended container

Return a useful classification where possible.

For example:

VALID_CONTAINMENT
SUSPECT_OVERLAP
ACTUAL_OVERLAP

At minimum, suppress the existing flood of false positives without hiding genuine layout problems.

Add tests using:
- card + text
- card + icon
- two overlapping cards
- two unrelated text boxes
- text overflowing container
- diagram objects overlapping

The validator should become trusted enough that the agent can act on its warnings.

==================================================
PRIORITY 6 — ADD ppt_batch_modify_shapes
==================================================

Create:

ppt_batch_modify_shapes

Purpose:

Allow multiple geometry edits in a single MCP call.

This addresses scenarios where the agent currently performs:

ppt_modify_shape × 10–20

for a single logical layout adjustment.

Input should support a list such as:

{
  "slide_number": 3,
  "operations": [
    {
      "shape_id": 13,
      "changes": {
        "height": 5.4
      }
    },
    {
      "shape_id": 14,
      "changes": {
        "y": 2.15
      }
    },
    {
      "shape_id": 15,
      "changes": {
        "y": 2.15
      }
    }
  ]
}

Support the existing geometry functionality already available in ppt_modify_shape.

Requirements:

- absolute and relative changes where currently supported
- multiple independent shape modifications in one call
- alignment/distribution where appropriate
- session-safe mutation
- concise results
- pre-validation of shape IDs
- safe failure behavior
- no partial corruption

Reuse internal geometry mutation functions rather than duplicating logic.

Add tests.

==================================================
SKILL UPDATE
==================================================

Update:

.agents/skills/powerpoint-editor/SKILL.md

Teach the agent to use the new capabilities.

Add explicit guidance:

1. For text-focused tasks:
   prefer ppt_inspect_text over full slide inspection.

2. For multiple text changes:
   prefer ppt_batch_modify_text.

3. For multiple geometry changes:
   prefer ppt_batch_modify_shapes.

4. Inspect once and reuse the result.

5. Do not write an external Python filtering script when the MCP can provide the filtered data directly.

6. Continue the existing workflow:

   inspect
   ↓
   modify
   ↓
   validate
   ↓
   render
   ↓
   correct if necessary
   ↓
   save

7. Do not pass presentation_path to mutation tools when an active session exists.

8. Preserve paragraph formatting by default during text replacement.

9. Treat validation warnings according to their classification rather than blindly reacting to all overlaps.

Explicitly encourage high-level batching while preserving precise deterministic operations.

==================================================
TOOL RESPONSE DESIGN
==================================================

Review all newly added tools for LLM usability.

Responses must be:

- structured
- concise
- actionable
- stable
- free of unnecessary raw data

Return:

- success/failure
- affected slide
- affected shape IDs
- concise change summary
- warnings/errors
- next-relevant state when useful

Do not return the entire slide state after every mutation.

For batch operations, summarize the batch rather than dumping every resulting shape property.

==================================================
REGRESSION TEST — ORIGINAL REAL-WORLD TASK
==================================================

After implementing all six priorities, rerun this exact task:

"On Slide 3. remove the Orchestration box, and its adjacent box. Extend the Client Configuration box to the bottom of the slide, and also add content and extend the box adjacent to the Client Configuration box."

Measure:

- total MCP tool calls
- number of inspection calls
- number of mutation calls
- number of validation calls
- number of render calls
- correction iterations

Compare against the previous run.

The goal is to reduce unnecessary mutation round trips substantially while producing the same visual result.

==================================================
REGRESSION TEST — TEXT TASK
==================================================

Rerun this exact task:

"Update slide 2. Clean up the text, and make it appropriate size for the slide. It is too small everywhere on this slide."

The intended agent workflow should now be approximately:

ppt_open
    ↓
ppt_inspect_text
    ↓
ppt_batch_modify_text
    ↓
ppt_validate_slide
    ↓
ppt_render_slide
    ↓
optional corrective ppt_batch_modify_text
    ↓
ppt_render_slide
    ↓
ppt_save / ppt_save_as

The agent should NOT need to:

- retrieve thousands of lines of slide JSON
- write a filtering Python script
- perform 20–30 individual text mutation MCP calls

The final rendered slide must still be visually verified.

==================================================
BENCHMARKING
==================================================

For both real-world regression tasks, record:

- total MCP calls
- calls by tool
- response size where measurable
- mutation calls
- render calls
- validation calls
- number of corrective iterations
- final result correctness

Do not optimize for tool-call count at the expense of correctness.

The objective is:

FEWER LOW-VALUE ROUND TRIPS
while preserving:
- precision
- safety
- visual fidelity
- deterministic editing
- session isolation

==================================================
OUT OF SCOPE FOR THIS ITERATION
==================================================

Do NOT implement yet:

- ppt_create_flow_diagram
- high-level diagram generation framework
- automatic typography scaling macro
- broad new diagram tools
- UI
- remote server
- LLM integration
- large MCP architecture rewrite

Those can be evaluated after these foundational agent-efficiency improvements are benchmarked.

==================================================
DELIVERABLE
==================================================

At completion provide:

1. Files changed.
2. New MCP tools.
3. Changes to existing tools.
4. Changes to SKILL.md.
5. Tests added.
6. Test results.
7. Before/after tool-call counts for the two regression tasks.
8. Before/after inspection output sizes where measurable.
9. Any remaining known limitations.

Do not stop at code changes.

Actually run the tests.

Actually run both PowerPoint regression tasks.

Actually render and verify the resulting slides.

Do not declare success based solely on unit tests.