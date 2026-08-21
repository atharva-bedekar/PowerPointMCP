# Build a Production-Quality PowerPoint Editing MCP Server

I want you to build a local MCP server that turns you, the Antigravity CLI agent, into a precise conversational PowerPoint editor.

The goal is NOT to generate presentations from scratch. The primary goal is to allow me to open an existing `.pptx` file and then make precise conversational edits to it:

- "Move the title 0.2 inches to the left."
- "Make slide 7 look like slide 6."
- "Change the font throughout this slide to Aptos."
- "Increase the spacing between these boxes."
- "Replace this diagram but preserve the existing style."
- "Align these three objects."
- "Make this slide match the visual language of the rest of the deck."
- "Move the image slightly down."
- "Make the title exactly the same size and position as slide 4."
- "Fix the overlap on slide 8."
- "Keep everything else unchanged."

The system must prioritize deterministic manipulation and visual verification over blindly asking an LLM to rewrite a PPTX.

---

# 1. Architecture

Build the project as a Python MCP server with this architecture:

    powerpoint-mcp/
    │
    ├── pyproject.toml
    ├── README.md
    ├── .gitignore
    │
    ├── src/
    │   └── powerpoint_mcp/
    │       ├── __init__.py
    │       ├── server.py
    │       │
    │       ├── models/
    │       │   ├── __init__.py
    │       │   ├── presentation.py
    │       │   ├── slide.py
    │       │   └── shape.py
    │       │
    │       ├── pptx/
    │       │   ├── __init__.py
    │       │   ├── inspector.py
    │       │   ├── editor.py
    │       │   ├── ooxml.py
    │       │   ├── geometry.py
    │       │   ├── styles.py
    │       │   └── relationships.py
    │       │
    │       ├── rendering/
    │       │   ├── __init__.py
    │       │   ├── renderer.py
    │       │   ├── image_diff.py
    │       │   └── visual_compare.py
    │       │
    │       ├── tools/
    │       │   ├── inspection.py
    │       │   ├── editing.py
    │       │   ├── rendering.py
    │       │   └── versioning.py
    │       │
    │       └── utils/
    │           ├── paths.py
    │           ├── validation.py
    │           └── logging.py
    │
    ├── tests/
    │   ├── fixtures/
    │   ├── test_inspection.py
    │   ├── test_editing.py
    │   ├── test_geometry.py
    │   ├── test_rendering.py
    │   └── test_mcp.py
    │
    ├── scripts/
    │   ├── render_pptx.py
    │   └── inspect_pptx.py
    │
    └── .agents/
        ├── skills/
        │   └── powerpoint-editor/
        │       └── SKILL.md
        │
        └── mcp_config.json

Use Python 3.10+.

Use the current stable MCP Python SDK rather than the legacy v1 API.

Use `uv` for dependency management if available.

---

# 2. Core principle

The LLM must NEVER blindly modify a PowerPoint file.

For every editing request, the intended workflow is:

    USER REQUEST
         ↓
    INSPECT
         ↓
    UNDERSTAND TARGET OBJECTS
         ↓
    MAKE MINIMAL DETERMINISTIC CHANGE
         ↓
    SAVE
         ↓
    RENDER AFFECTED SLIDE(S)
         ↓
    VISUAL VERIFICATION
         ↓
    IF NECESSARY → CORRECT
         ↓
    FINAL SAVE

The agent should prefer modifying existing objects over recreating them.

Preserve whenever possible:

- existing theme
- slide master
- layouts
- fonts
- colors
- shape styles
- animations
- relationships
- object IDs
- z-order
- speaker notes
- hyperlinks
- embedded objects
- charts
- tables
- images
- existing XML not directly related to the requested change

Do not reconstruct an entire slide unless explicitly necessary.

---

# 3. The ten core capabilities

Expose these capabilities through MCP tools.

## 3.1 inspect_presentation

Tool:

    ppt_inspect_presentation

Input:

    presentation_path: string

Return structured information containing:

- presentation dimensions
- slide count
- theme information
- available layouts
- slide titles
- slide sizes
- notes availability
- basic metadata

Do not return enormous raw XML.

Return concise structured information useful to an LLM.

---

## 3.2 inspect_slide

Tool:

    ppt_inspect_slide

Inputs:

    presentation_path: string
    slide_number: integer

Return every meaningful object on the slide.

For each shape include:

- shape ID
- shape name
- semantic type
- x
- y
- width
- height
- rotation
- z-order
- text
- text frame information
- font family
- font size
- bold
- italic
- color
- fill
- line
- alignment
- margins
- paragraph spacing
- image metadata
- table metadata
- chart metadata
- group membership

Coordinates must be returned in inches.

Also provide semantic hints such as:

    role: title
    role: subtitle
    role: body
    role: image
    role: diagram
    role: footer
    role: unknown

Infer these roles conservatively.

The output should make it easy for the agent to say:

    "Shape 14 is the title."

rather than having to reason over raw XML.

---

## 3.3 inspect_shape

Tool:

    ppt_inspect_shape

Inputs:

    presentation_path: string
    slide_number: integer
    shape_id: integer

Return detailed information about exactly one shape.

Include all editable properties exposed by python-pptx and relevant OOXML properties when available.

---

## 3.4 modify_shape

Tool:

    ppt_modify_shape

Inputs should support:

- presentation_path
- slide_number
- shape_id
- x
- y
- width
- height
- rotation
- z_order

All parameters except identifiers should be optional.

Only modify explicitly provided properties.

Do not silently modify unrelated properties.

Support alignment helpers:

- align_left
- align_center
- align_right
- align_top
- align_middle
- align_bottom
- distribute_horizontal
- distribute_vertical

where practical.

---

## 3.5 modify_text

Tool:

    ppt_modify_text

Inputs should support:

- presentation_path
- slide_number
- shape_id
- text
- font_family
- font_size
- bold
- italic
- underline
- color
- alignment
- paragraph_spacing
- line_spacing
- margins

Support modifying an entire text box as well as individual text runs where practical.

Do not destroy existing rich text formatting unnecessarily.

If only one word or phrase changes, preserve the formatting of surrounding runs.

---

## 3.6 copy / move / resize / delete

Create tools for:

    ppt_copy_shape
    ppt_move_shape
    ppt_resize_shape
    ppt_delete_shape

Moving and resizing should be deterministic.

For copy operations preserve:

- style
- formatting
- image
- relationships where possible

Do not use screenshot-based recreation for ordinary objects.

---

## 3.7 OOXML manipulation

Create a lower-level tool:

    ppt_modify_ooxml

This is an escape hatch for functionality unavailable through python-pptx.

It must NOT allow arbitrary uncontrolled modification by default.

Implement helper functions in:

    pptx/ooxml.py

Use them for cases such as:

- unsupported PowerPoint properties
- advanced shape properties
- unsupported formatting
- XML attributes not exposed by python-pptx

Prefer purpose-built OOXML helpers over arbitrary XML string replacement.

Preserve namespaces and relationships correctly.

Create backups before OOXML modifications.

---

## 3.8 render_slide / render_presentation

Tools:

    ppt_render_slide
    ppt_render_presentation

Render slides to PNG.

Primary rendering strategy on Windows:

1. Detect Microsoft PowerPoint if installed.
2. Prefer PowerPoint COM automation for highest fidelity.
3. Fall back to LibreOffice headless rendering if PowerPoint is unavailable.
4. Clearly report which renderer was used.

Do NOT assume LibreOffice and PowerPoint render identically.

Rendering output should go into a temporary working directory.

For example:

    .ppt-agent/
        renders/
            presentation_20260821/
                slide-01.png
                slide-02.png

Do not clutter the source directory.

---

# 4. Visual verification

This is a critical part of the system.

Create:

    ppt_compare_slides

Inputs:

    presentation_path
    slide_a
    slide_b

It should support comparing:

- geometry
- text
- dimensions
- alignment
- relative positioning
- visual appearance

Also create:

    ppt_visual_diff

Inputs:

    before_image
    after_image

Generate:

- difference image
- changed bounding regions
- basic similarity metrics

The agent should be able to use this after edits.

---

# 5. Reference-slide workflow

The system must make this workflow easy:

    "Make slide 8 look like slide 6."

The agent should:

1. Inspect slide 6.
2. Inspect slide 8.
3. Identify corresponding semantic objects.
4. Compare geometry and formatting.
5. Determine which properties differ.
6. Apply the minimum required changes.
7. Render slide 8.
8. Compare slide 8 against slide 6.
9. Correct obvious mismatches.
10. Render again.
11. Report what was changed.

Do not simply duplicate slide 6.

The content of slide 8 must remain intact unless the user explicitly asks for content replacement.

---

# 6. Semantic matching

Implement a utility:

    match_shapes(slide_a, slide_b)

It should attempt to identify corresponding objects using:

1. explicit shape names
2. semantic role
3. text similarity
4. shape type
5. relative position
6. dimensions
7. group membership

Return confidence scores.

Example:

    {
      "source_shape": 12,
      "target_shape": 27,
      "confidence": 0.94,
      "reason": [
        "same semantic role: title",
        "similar text",
        "same relative position"
      ]
    }

The LLM should be able to use this rather than manually guessing correspondence.

---

# 7. Geometry utilities

Create robust geometry functions.

Support:

- distance
- bounding boxes
- intersection
- overlap detection
- alignment
- centering
- equal sizing
- distribution
- margins
- relative positioning

Examples:

    align_shapes_left(...)
    align_shapes_center(...)
    distribute_shapes_horizontal(...)
    equalize_width(...)
    equalize_height(...)
    detect_overlaps(...)

Return coordinates in inches.

Use PowerPoint's native EMU representation internally where appropriate.

Do not repeatedly convert between floating point inches and EMUs in ways that introduce cumulative rounding errors.

---

# 8. Detect common PowerPoint problems

Create:

    ppt_validate_slide

It should detect:

- overlapping objects
- objects outside slide boundaries
- text overflowing text boxes
- suspiciously tiny fonts
- inconsistent title positions
- inconsistent margins
- duplicate objects
- extreme rotations
- accidental off-slide objects
- unusually large images
- missing fonts where detectable

Return warnings rather than automatically changing things.

Example:

    WARNING:
    Shape 32 overlaps Shape 34 by 0.21 inches.

    WARNING:
    Shape 17 extends 0.12 inches beyond the right slide boundary.

---

# 9. Versioning / safety

Every modification must create a backup before modifying the source file.

Implement:

    ppt_create_backup

Use timestamped versions.

Example:

    presentation.pptx
    presentation.backup-20260821-111530.pptx

Also support:

    ppt_save_as

Never overwrite the original unless the user explicitly requests it.

The MCP server should maintain a session working copy where possible.

---

# 10. Working-copy model

When a presentation is first opened, create a working copy.

For example:

    source/
        DaaS.pptx

    .ppt-agent/
        sessions/
            <session-id>/
                working.pptx
                backups/
                renders/
                diffs/
                metadata.json

All edits should operate on the working copy.

The user should eventually be able to say:

    "Save this."

which copies the working presentation to a specified output path.

Implement:

    ppt_open
    ppt_save
    ppt_save_as
    ppt_revert

If a persistent session model is difficult for the first version, implement a simpler explicit `working_path`, but design the architecture so sessions can be added later.

---

# 11. MCP tool design

Keep the MCP tool surface reasonably small.

Do not expose dozens of tiny tools if several operations can be represented safely by one structured tool.

The initial exposed tool set should be approximately:

    ppt_open
    ppt_inspect_presentation
    ppt_inspect_slide
    ppt_inspect_shape
    ppt_modify_shape
    ppt_modify_text
    ppt_copy_shape
    ppt_move_shape
    ppt_resize_shape
    ppt_delete_shape
    ppt_modify_ooxml
    ppt_validate_slide
    ppt_render_slide
    ppt_render_presentation
    ppt_compare_slides
    ppt_visual_diff
    ppt_save
    ppt_save_as
    ppt_revert

Add tools only when they genuinely improve agent reliability.

Every tool must have excellent descriptions because the LLM will use the MCP schema to decide when and how to invoke it.

Use structured return values.

Do not return giant blobs of unstructured text.

---

# 12. Resources

Expose useful MCP resources if appropriate.

At minimum consider:

    ppt://current/presentation
    ppt://current/slide/{slide_number}
    ppt://current/slide/{slide_number}/render

The resource representation should be concise.

Do not automatically load the entire presentation into context.

---

# 13. Antigravity PowerPoint skill

Create:

    .agents/skills/powerpoint-editor/SKILL.md

Use the current Antigravity skill format with YAML frontmatter.

The skill should teach the agent the following workflow.

## PowerPoint editing rules

1. Always inspect before editing.
2. Identify objects semantically.
3. Make the smallest possible change.
4. Preserve existing styles.
5. Never recreate an object when it can be modified.
6. Never rebuild an entire slide unless necessary.
7. After visual changes, render the affected slide.
8. Inspect the rendered result.
9. If the result is wrong, make another correction.
10. Save only after verification.
11. Prefer exact geometric operations over vague visual changes.
12. When the user references another slide, inspect that slide first.
13. When matching slides, preserve target content while copying relevant style/layout characteristics.
14. Do not silently alter unrelated slides.
15. When uncertain which object the user means, inspect the slide rather than guessing.

The skill should contain a decision tree.

Example:

    User asks to change text
        ↓
    inspect slide
        ↓
    identify text shape
        ↓
    modify text
        ↓
    render if visual formatting may change
        ↓
    verify

    User asks to move/resize something
        ↓
    inspect slide
        ↓
    identify shape
        ↓
    modify geometry
        ↓
    render
        ↓
    verify

    User asks "make slide A like slide B"
        ↓
    inspect both
        ↓
    match shapes
        ↓
    compare
        ↓
    modify target
        ↓
    render
        ↓
    compare again
        ↓
    correct

---

# 14. Agent behavior

The skill must strongly discourage the agent from doing this:

    read PPT
    generate entirely new PPT
    replace old PPT

Instead it should do:

    inspect
    reason
    modify
    render
    verify

The agent should also avoid excessive MCP calls.

Batch inspection where possible.

For example, do not call `ppt_inspect_shape` 20 times if `ppt_inspect_slide` already contains everything necessary.

Use the least number of tool calls required to safely perform the edit.

---

# 15. Rendering implementation

Implement the rendering abstraction:

    class Renderer:
        def render_presentation(...)
        def render_slide(...)

Implement:

    PowerPointRenderer
    LibreOfficeRenderer

PowerPointRenderer should use COM automation on Windows if PowerPoint is installed.

Requirements:

- run PowerPoint invisibly
- open presentation
- export slides
- close presentation
- close PowerPoint
- clean up COM objects
- never leave orphaned PowerPoint processes

LibreOffice fallback should use headless mode.

Detect the available renderer automatically.

Return renderer information in the MCP response.

---

# 16. Visual comparison

Do not make the visual comparison system depend entirely on an LLM.

Use deterministic image processing for:

- pixel difference
- changed regions
- bounding boxes
- image dimensions
- similarity percentage

If useful, optionally expose the resulting images to the agent as MCP image content.

The goal is to let the agent visually inspect the result while also having deterministic measurements.

---

# 17. Tests

Build comprehensive automated tests.

At minimum test:

## Inspection

- presentation inspection
- slide inspection
- shape inspection

## Geometry

- move
- resize
- alignment
- distribution
- overlap detection
- boundary detection

## Text

- text replacement
- font modification
- preserving unrelated formatting
- rich text preservation

## Editing

- copy
- delete
- z-order
- image manipulation

## OOXML

- unsupported property modification
- namespace preservation
- relationship preservation

## Rendering

- PowerPoint renderer detection
- LibreOffice fallback
- rendered file existence

## MCP

Test every MCP tool through an in-memory MCP client.

Follow the current MCP SDK testing approach rather than spawning a subprocess for every unit test.

---

# 18. Example test presentation

Create a small test presentation programmatically containing:

Slide 1:
- title
- subtitle
- three boxes
- image placeholder

Slide 2:
- title
- two-column layout
- diagram
- footer

Slide 3:
- title
- several overlapping objects intentionally

Use this presentation to test the complete workflow.

---

# 19. CLI utilities

Create:

    scripts/inspect_pptx.py

Example:

    python scripts/inspect_pptx.py presentation.pptx

Create:

    scripts/render_pptx.py

Example:

    python scripts/render_pptx.py presentation.pptx --output renders/

These should be useful for debugging the MCP independently of Antigravity.

---

# 20. Configuration

Create a configuration system supporting:

    PPT_RENDERER=auto
    PPT_WORKSPACE_DIR=.ppt-agent
    PPT_BACKUP_ENABLED=true
    PPT_DEFAULT_OUTPUT_DIR=./output

Do not hard-code paths.

Use environment variables where appropriate.

---

# 21. Antigravity MCP configuration

Create:

    .agents/mcp_config.json

Configure the local server using stdio.

The configuration should launch the MCP server using the project's Python environment / uv environment.

Do not use a remote HTTP server for the initial version.

The intended architecture is:

    Antigravity CLI
        |
        | stdio
        v
    powerpoint-mcp
        |
        +-- python-pptx
        +-- OOXML
        +-- PowerPoint COM
        +-- LibreOffice
        +-- image processing

Antigravity's current CLI supports workspace MCP configuration through `.agents/mcp_config.json`, so make the project self-contained and launchable from its root.

---

# 22. README

Write a complete README explaining:

1. What this project does.
2. Architecture.
3. Requirements.
4. Installation.
5. Python/uv setup.
6. Optional Microsoft PowerPoint requirement.
7. Optional LibreOffice requirement.
8. How to run the MCP server.
9. How to configure Antigravity.
10. How to verify the server.
11. Example conversational commands.
12. Troubleshooting.

Include examples such as:

    "Inspect slide 5."

    "Move the title 0.2 inches left."

    "Make all three boxes the same width."

    "Align these boxes vertically."

    "Make slide 8 match the layout of slide 6 but preserve slide 8's content."

    "Render slide 8."

    "Check whether anything overlaps on slide 8."

---

# 23. Important implementation constraints

Do NOT:

- build a web UI
- build a presentation generator
- use an LLM API inside the MCP server
- send the PPTX to an external service
- rasterize slides and rebuild them from screenshots
- replace the entire presentation when making a small change
- depend on OpenAI/Claude/Gemini APIs
- require a cloud backend

This MCP is a deterministic local editing engine.

The LLM is Antigravity itself.

---

# 24. Dependency choices

Prefer:

- Python 3.10+
- current stable MCP Python SDK
- python-pptx
- Pillow
- numpy
- lxml if necessary for OOXML
- pywin32 on Windows for PowerPoint COM
- pytest
- uv

Use the smallest dependency set that works.

Before adding a dependency, determine whether the functionality can be implemented with the existing stack.

---

# 25. Error handling

MCP errors must be useful to an LLM.

Bad:

    "Exception occurred."

Good:

    {
      "success": false,
      "error_type": "ShapeNotFound",
      "message": "Shape ID 17 does not exist on slide 5.",
      "available_shapes": [12, 13, 14, 19]
    }

When an operation fails, provide enough structured information for the agent to recover without guessing.

Never silently swallow errors.

---

# 26. Idempotency

Where possible, operations should be idempotent.

For example:

    set shape 12 x = 1.2

is preferable to:

    move shape 12 left by 0.2

because repeated execution should not accumulate unintended changes.

Support both absolute and relative operations where useful.

---

# 27. Precision

PowerPoint coordinates ultimately use EMUs.

Use integer EMUs internally where practical.

Expose inches to the LLM.

Example:

    1 inch = 914400 EMU

Do not introduce unnecessary rounding.

When reporting coordinates, use enough precision to make small adjustments possible.

For example:

    x = 1.2375 inches

rather than:

    x = 1.2 inches

---

# 28. Minimal-diff philosophy

Every edit should produce the smallest possible change to the PPTX.

For example:

User:

    "Move the title 0.1 inches left."

The agent should NOT:

- recreate the title
- change its font
- change its size
- change its color
- change the slide layout

It should only change:

    title.left -= 0.1 inches

This principle is extremely important.

---

# 29. Final integration test

After implementing everything, perform this end-to-end test.

Create/open a sample presentation.

Then simulate these user requests:

### Test 1

    "Inspect slide 1."

Expected:
Structured slide description.

### Test 2

    "Move the title 0.2 inches to the left."

Expected:
Only title X coordinate changes.

### Test 3

    "Make the three boxes the same width and distribute them evenly."

Expected:
Geometry changes only.

### Test 4

    "Make slide 2 match the layout of slide 1 but preserve slide 2's content."

Expected:
The agent inspects both slides, matches objects, modifies slide 2, renders it, and verifies it.

### Test 5

    "Find anything overlapping on slide 3."

Expected:
Overlap report.

### Test 6

    "Render slide 2."

Expected:
PNG output.

### Test 7

    "Save the result as output/final.pptx."

Expected:
New PPTX created without modifying the original.

---

# 30. Do not stop at scaffolding

I want an actually working implementation.

Do not merely create:

- empty classes
- TODO comments
- placeholder MCP tools
- fake rendering
- mocked PowerPoint editing

Implement the real functionality.

If some PowerPoint feature cannot be supported through python-pptx, implement an OOXML fallback where reasonable.

If PowerPoint COM is unavailable, implement LibreOffice fallback and clearly report limitations.

---

# 31. Build order

Implement in this order:

1. Project setup / dependencies
2. PPTX loading and inspection
3. Geometry utilities
4. Basic editing
5. Text editing
6. OOXML utilities
7. Rendering
8. Visual comparison
9. Validation
10. MCP server
11. Tests
12. Antigravity skill
13. Antigravity MCP configuration
14. End-to-end test
15. README

After each major phase, run tests.

Do not wait until the end to discover architectural problems.

---

# 32. Important: inspect the environment first

Before writing code:

1. Determine OS.
2. Determine Python version.
3. Determine whether `uv` is installed.
4. Determine whether Microsoft PowerPoint is installed.
5. Determine whether LibreOffice is installed.
6. Determine whether Node/npx is installed.
7. Determine the current Antigravity workspace structure.
8. Check whether an existing `.agents` directory exists.
9. Check whether existing MCP configuration exists.
10. Check whether there are existing skills or agents that must not be overwritten.

Do not overwrite existing user configuration.

If `.agents/mcp_config.json` already exists, merge the new server entry rather than replacing the file.

If `.agents/skills` already exists, add the new skill without deleting anything.

---

# 33. Use current documentation when necessary

If you encounter uncertainty about:

- current MCP Python SDK APIs
- current Antigravity MCP configuration
- current Antigravity skill format
- current Antigravity agent format

use the official documentation rather than relying on memory.

Prefer official MCP and Antigravity documentation.

---

# 34. Completion criteria

Consider the task complete only when:

- MCP server starts successfully.
- Antigravity can discover its tools.
- `ppt_inspect_presentation` works.
- `ppt_inspect_slide` works.
- `ppt_inspect_shape` works.
- geometry editing works.
- text editing works.
- copy/delete works.
- OOXML fallback exists.
- PowerPoint/LibreOffice rendering works.
- visual comparison works.
- slide validation works.
- backups work.
- save-as works.
- automated tests pass.
- Antigravity skill exists.
- Antigravity MCP configuration exists.
- README exists.
- an end-to-end edit has been performed successfully.

At the end, give me:

1. A concise architecture summary.
2. The exact command to install dependencies.
3. The exact command to test the MCP server.
4. The exact Antigravity command/workflow to load it.
5. The location of the skill.
6. The location of the MCP configuration.
7. A list of supported MCP tools.
8. Any environment limitations you discovered.
9. The end-to-end test result.

Do not stop after creating files. Actually run the tests and fix failures.