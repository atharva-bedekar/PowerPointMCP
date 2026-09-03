\# PowerPoint MCP v1.2 — Core Editing Completeness \& Reliability



\## Objective



Harden the PowerPoint MCP and eliminate the major reasons agents currently fall back to custom

python-pptx scripting.



v1.2 should focus on making the MCP a sufficiently complete PowerPoint editing interface for

common conversational editing tasks.



The primary goals are:



1\. Fix correctness bugs in existing mutation APIs.

2\. Make rendering reliable on Windows.

3\. Add native image insertion/replacement.

4\. Add first-class table inspection and mutation.

5\. Add batch operations for tables and media where useful.

6\. Extend validation to understand tables and inserted media.

7\. Reduce the need for custom python-pptx scripts.



Do NOT spend this version primarily on new high-level AI/design abstractions. Those should come

after the core PowerPoint primitives are reliable.



\---



\# P0 — Correctness



\## 1. Fix `ppt\_modify\_shape` partial-update bug



\### Current bug



Calling:



&#x20;   ppt\_modify\_shape(shape\_id=2, slide\_number=2, height=0.8)



can reset omitted properties such as width to `0.0`.



This is unacceptable because it allows an apparently harmless partial mutation to destroy

existing geometry.



\### Required behavior



Mutation arguments must be PATCH semantics.



If the caller provides:



&#x20;   height=0.8



then ONLY height changes.



All other properties must retain their existing values:



\- left

\- top

\- width

\- height

\- rotation

\- etc.



For example:



&#x20;   current:

&#x20;       left=1.2

&#x20;       top=2.0

&#x20;       width=8.0

&#x20;       height=1.0



&#x20;   modify:

&#x20;       height=0.8



&#x20;   result:

&#x20;       left=1.2

&#x20;       top=2.0

&#x20;       width=8.0

&#x20;       height=0.8



Never convert omitted values into zero.



\### Implementation



Inspect the schema and implementation of `ppt\_modify\_shape`.



Distinguish:



&#x20;   argument omitted



from:



&#x20;   argument explicitly supplied as 0



These must not be treated identically.



Use optional/nullable fields appropriately.



Before mutation, obtain the existing shape geometry and construct the final geometry:



&#x20;   final = existing.copy()

&#x20;   final.update(only\_explicitly\_supplied\_fields)



Do not change the public API unnecessarily.



\### Tests



Add regression tests for:



\- height only

\- width only

\- left only

\- top only

\- width + height

\- left + top

\- explicit width=0 rejection/handling

\- multiple properties

\- batch shape modification



The regression test must prove that omitted dimensions remain unchanged.



\---



\# P0 — Rendering



\## 2. Make `ppt\_render\_slide` reliable on Windows



Rendering is a core capability, not an optional utility.



The MCP must support a Windows environment where:



\- Microsoft PowerPoint is installed

\- LibreOffice is not installed



\### Preferred backend



Use Microsoft PowerPoint COM when available.



Conceptually:



&#x20;   PowerPoint.Application

&#x20;       -> Open presentation

&#x20;       -> Export/render requested slide

&#x20;       -> Close presentation

&#x20;       -> Release COM objects



\### Fallback hierarchy



Use an explicit backend strategy:



&#x20;   1. PowerPoint COM

&#x20;   2. LibreOffice

&#x20;   3. Existing fallback renderer, if present



Do not assume LibreOffice exists.



\### COM requirements



Rendering must NEVER leave POWERPNT.EXE holding the presentation open.



Ensure:



\- Presentation.Close()

\- Application.Quit() where appropriate

\- COM references released

\- temporary objects released

\- retry/cleanup on exceptions

\- no lingering presentation file handles



Use `try/finally`.



\### Important



Do not terminate all POWERPNT.EXE processes as the normal cleanup strategy.



Do not kill the user's unrelated PowerPoint sessions.



Only clean up COM processes/resources owned by the MCP where possible.



\### Backend selection



Expose enough diagnostic information for the agent/logs to know:



&#x20;   renderer: powerpoint\_com



or:



&#x20;   renderer: libreoffice



or:



&#x20;   renderer: fallback



Do not make the LLM reason about renderer selection unless necessary.



\---



\# P0 — Images / Media



\## 3. Add `ppt\_add\_picture`



Add first-class image insertion.



Interface:



&#x20;   ppt\_add\_picture(

&#x20;       slide\_number,

&#x20;       image\_path,

&#x20;       left,

&#x20;       top,

&#x20;       width=None,

&#x20;       height=None,

&#x20;       preserve\_aspect\_ratio=True

&#x20;   )



\### Requirements



Support:



\- PNG

\- JPEG/JPG

\- BMP if supported by PowerPoint

\- transparent PNGs



If width and height are omitted:



\- use native image dimensions where practical

\- otherwise use a reasonable PowerPoint size



If only width is supplied:



\- calculate height from aspect ratio



If only height is supplied:



\- calculate width from aspect ratio



If both are supplied:



\- honor explicit dimensions unless preserve\_aspect\_ratio requires adjustment



Return:



\- shape\_id

\- final geometry

\- image dimensions

\- slide number



\### Session safety



The inserted image must be added to:



&#x20;   session.working\_path



when an active session exists.



Never bypass the active working presentation.



\---



\# 4. Add `ppt\_replace\_picture`



Add:



&#x20;   ppt\_replace\_picture(

&#x20;       slide\_number,

&#x20;       shape\_id,

&#x20;       image\_path,

&#x20;       preserve\_geometry=True

&#x20;   )



When `preserve\_geometry=true`:



\- replace image content

\- preserve left

\- preserve top

\- preserve width

\- preserve height

\- preserve rotation where possible



This is especially useful for image placeholders.



\---



\# P0/P1 — Tables



\## 5. Add table inspection at cell level



Extend the existing inspection system.



Add:



&#x20;   ppt\_inspect\_table



or equivalent functionality in the existing inspection API.



Return:



\- table shape ID

\- row count

\- column count

\- table bounding box

\- column widths

\- row heights

\- cell text

\- cell coordinates

\- font properties

\- paragraph alignment

\- margins

\- fill

\- borders

\- merge state



Do NOT dump enormous XML.



Default output must be compact and agent-friendly.



Example:



&#x20;   Table 12

&#x20;   5 rows × 3 columns

&#x20;   bbox: ...



&#x20;   R1:

&#x20;     C1: "Application"

&#x20;     C2: "Owner"

&#x20;     C3: "Status"



&#x20;   R2:

&#x20;     C1: "Portal"

&#x20;     C2: "Platform"

&#x20;     C3: "Active"



Allow a `detail=full` mode for deeper diagnostics.



\---



\# 6. Add `ppt\_batch\_modify\_table\_cells`



This should be the primary table text mutation API.



Example:



&#x20;   ppt\_batch\_modify\_table\_cells(

&#x20;       slide\_number=3,

&#x20;       table\_shape\_id=12,

&#x20;       mutations=\[

&#x20;           {

&#x20;               "row": 0,

&#x20;               "column": 0,

&#x20;               "text": "Application"

&#x20;           },

&#x20;           {

&#x20;               "row": 1,

&#x20;               "column": 2,

&#x20;               "text": "Active",

&#x20;               "font\_size": 12,

&#x20;               "bold": true

&#x20;           }

&#x20;       ]

&#x20;   )



The operation should be atomic.



Do not require one MCP call per cell.



\---



\# 7. Add table geometry operations



Add:



&#x20;   ppt\_set\_table\_geometry



Support:



\- individual column widths

\- individual row heights

\- total table geometry

\- optional auto-distribution



Example:



&#x20;   ppt\_set\_table\_geometry(

&#x20;       slide\_number=3,

&#x20;       table\_shape\_id=12,

&#x20;       column\_widths=\[2.2, 3.0, 1.8],

&#x20;       row\_heights=\[0.5, 0.6, 0.6]

&#x20;   )



Do not require the caller to specify values that are unchanged.



Again, use PATCH semantics.



\---



\# 8. Add `ppt\_style\_table`



Support common formatting operations.



At minimum:



\- cell fill

\- font family

\- font size

\- bold

\- italic

\- font color

\- horizontal alignment

\- vertical alignment

\- cell margins

\- borders



Support applying formatting to:



\- individual cells

\- rectangular ranges

\- rows

\- columns

\- entire table



Example:



&#x20;   ppt\_style\_table(

&#x20;       slide\_number=3,

&#x20;       table\_shape\_id=12,

&#x20;       range="0:0-0:2",

&#x20;       style={

&#x20;           "bold": true,

&#x20;           "font\_size": 12,

&#x20;           "horizontal\_alignment": "center"

&#x20;       }

&#x20;   )



\---



\# 9. Add `ppt\_merge\_table\_cells`



Support:



&#x20;   ppt\_merge\_table\_cells(

&#x20;       slide\_number,

&#x20;       table\_shape\_id,

&#x20;       start\_row,

&#x20;       start\_column,

&#x20;       end\_row,

&#x20;       end\_column

&#x20;   )



Validate ranges before mutation.



Return the resulting table structure.



\---



\# P1 — Batch Efficiency



\## 10. Add `ppt\_batch\_modify\_tables`



For tasks involving multiple tables on one or multiple slides, provide a single transaction.



Example:



&#x20;   ppt\_batch\_modify\_tables(

&#x20;       operations=\[

&#x20;           {

&#x20;               "slide": 3,

&#x20;               "table": 12,

&#x20;               "cells": \[...]

&#x20;           },

&#x20;           {

&#x20;               "slide": 6,

&#x20;               "table": 18,

&#x20;               "cells": \[...]

&#x20;           }

&#x20;       ]

&#x20;   )



This is important because table-heavy presentations can otherwise generate

dozens of round trips.



The operation must preserve session safety and be atomic where practical.



\---



\# P1 — Validation



\## 11. Extend `ppt\_validate\_slide` for tables



Add table-specific validation.



Detect:



\### Boundary overflow



&#x20;   table.bottom > slide.height



\### Cell content overflow



Detect text that cannot fit in the cell.



Prefer actual PowerPoint/COM measurements where practical.



Do not rely solely on character counts.



\### Row-height problems



Detect rows whose content requires substantially more space than the assigned height.



\### Column-width problems



Detect severe clipping or wrapping problems.



\### Table collision



Treat the table as a structural object while still detecting genuine collisions

with unrelated objects.



Do not report every cell as an independent overlap.



Example output:



&#x20;   TABLE-01

&#x20;   Table 12 extends 0.31" below slide boundary



&#x20;   TABLE-02

&#x20;   Table 12 row 4 has insufficient height for its text



\---



\# P1 — Batch Rendering



\## 12. Add `ppt\_render\_slides`



Support:



&#x20;   ppt\_render\_slides(

&#x20;       slide\_numbers=\[3,4,5,6]

&#x20;   )



Return references to all generated images.



Use the existing high-fidelity renderer.



Do not open/close PowerPoint separately for every slide.



Prefer:



&#x20;   open presentation once

&#x20;   render all requested slides

&#x20;   close presentation once



This should substantially reduce COM startup overhead.



\---



\# P1 — Batch Validation



\## 13. Add `ppt\_validate\_slides`



Support:



&#x20;   ppt\_validate\_slides(

&#x20;       slide\_numbers=\[3,4,5,6]

&#x20;   )



Return compact per-slide summaries:



&#x20;   Slide 3

&#x20;     errors: 0

&#x20;     warnings: 1



&#x20;   Slide 4

&#x20;     errors: 0

&#x20;     warnings: 0



Only return detailed findings when requested.



\---



\# P2 — Semantic Components



Do not implement the full semantic component system before the above capabilities are stable.



However, design the architecture so v1.2 can support it later.



The semantic layer should eventually provide:



&#x20;   header

&#x20;   footer

&#x20;   stepper

&#x20;   card

&#x20;   card\_list

&#x20;   title\_block

&#x20;   metric\_group



Do not hard-code these into the table/media implementation.



The shape-level primitives should remain independent.



\---



\# P2 — Cross-Slide Synchronization



After core primitives are complete, consider:



&#x20;   ppt\_compare\_slides

&#x20;   ppt\_sync\_component

&#x20;   ppt\_sync\_slide\_chrome

&#x20;   ppt\_sync\_layout



These should be implemented on top of the lower-level primitives rather than replacing them.



\---



\# Skill / Agent Instructions



Update the PowerPoint editing skill with the following priority order:



&#x20;   1. Use high-level MCP operation when available.

&#x20;   2. Use batch MCP operation.

&#x20;   3. Use individual MCP primitive.

&#x20;   4. Use custom python-pptx only when the MCP genuinely lacks the required capability.



Specifically:



\### Images



Never write custom python-pptx merely to insert an image if `ppt\_add\_picture`

is available.



\### Tables



Never write custom python-pptx merely to:



\- edit table text

\- change font sizes

\- change cell formatting

\- resize columns

\- resize rows

\- merge cells



if the corresponding MCP tools exist.



\### Geometry



Prefer PATCH-style mutations.



If only one geometry property needs to change, only provide that property.



\### Rendering



Always render after substantial visual modifications.



Prefer batch rendering when inspecting multiple slides.



\### Validation



Validate after substantial layout modifications.



Prefer batch validation for multi-slide tasks.



\---



\# Testing



Add automated regression tests for all new functionality.



\## Shape mutation



Test:



&#x20;   height only

&#x20;   width only

&#x20;   x only

&#x20;   y only

&#x20;   multiple properties

&#x20;   explicit zero

&#x20;   batch modifications



\## Rendering



Test:



&#x20;   PowerPoint installed + LibreOffice absent

&#x20;   LibreOffice installed + PowerPoint unavailable

&#x20;   rendering multiple slides

&#x20;   renderer cleanup

&#x20;   exception cleanup

&#x20;   repeated rendering



Ensure no unintended POWERPNT.EXE process/file lock remains.



\## Images



Test:



&#x20;   add PNG

&#x20;   add JPEG

&#x20;   aspect ratio preservation

&#x20;   explicit dimensions

&#x20;   replace picture

&#x20;   preserve geometry

&#x20;   active session behavior



\## Tables



Test:



&#x20;   inspect cells

&#x20;   modify cells

&#x20;   batch cell modification

&#x20;   style cell

&#x20;   style range

&#x20;   change column widths

&#x20;   change row heights

&#x20;   merge cells

&#x20;   multi-table batch operations

&#x20;   active session behavior



\## Validation



Test:



&#x20;   table boundary overflow

&#x20;   table text overflow

&#x20;   valid table containment

&#x20;   genuine table collision

&#x20;   normal text-in-cell containment



\---



\# Acceptance Test



The following real-world task should be executable almost entirely through MCP:



&#x20;   "Update slides 3, 4 and 6. Populate the tables with this content,

&#x20;   make the headers bold, resize the columns so the text fits,

&#x20;   insert this architecture diagram into slide 4,

&#x20;   make sure nothing overflows the slide,

&#x20;   render the affected slides and fix anything that looks wrong."



Expected workflow:



&#x20;   ppt\_open

&#x20;       ↓

&#x20;   ppt\_inspect\_table / ppt\_inspect\_slide

&#x20;       ↓

&#x20;   ppt\_batch\_modify\_table\_cells

&#x20;       ↓

&#x20;   ppt\_style\_table

&#x20;       ↓

&#x20;   ppt\_set\_table\_geometry

&#x20;       ↓

&#x20;   ppt\_add\_picture

&#x20;       ↓

&#x20;   ppt\_validate\_slides

&#x20;       ↓

&#x20;   ppt\_render\_slides

&#x20;       ↓

&#x20;   corrective MCP operation if required

&#x20;       ↓

&#x20;   ppt\_validate\_slides

&#x20;       ↓

&#x20;   ppt\_render\_slides

&#x20;       ↓

&#x20;   ppt\_save\_as



The agent should NOT need to create a custom python-pptx script for:



\- table population

\- table styling

\- table geometry

\- image insertion

\- rendering

\- validation



\---



\# v1.2 Success Metrics



Target:



\### Reliability



\- 0 known partial-update geometry corruption bugs

\- 0 known session-buffer bypasses

\- rendering works without LibreOffice when PowerPoint is available

\- rendering does not leave MCP-owned PowerPoint locks



\### Efficiency



For a table-heavy slide:



&#x20;   Before:

&#x20;   20–50+ individual operations / custom scripting



&#x20;   Target:

&#x20;   3–8 MCP calls



For image insertion:



&#x20;   Before:

&#x20;   custom python-pptx



&#x20;   Target:

&#x20;   1 MCP call



For multi-slide rendering:



&#x20;   Before:

&#x20;   one renderer invocation per slide



&#x20;   Target:

&#x20;   1 batch render call



\### Capability



The MCP should natively support the majority of ordinary operations involving:



\- text

\- shapes

\- images

\- tables

\- geometry

\- rendering

\- validation



before v1.2 is considered complete.



\---



\# Design Principle



v1.2 is about eliminating the "escape hatch" problem.



The MCP should first become a reliable PowerPoint manipulation substrate.



Only after that should increasingly sophisticated abstractions be added.



The desired architecture is:



&#x20;   Conversational intent

&#x20;           ↓

&#x20;   Agent reasoning

&#x20;           ↓

&#x20;   Semantic operations (future)

&#x20;           ↓

&#x20;   Batch operations

&#x20;           ↓

&#x20;   Core PowerPoint primitives

&#x20;      ┌────┼────┬────┐

&#x20;      ↓    ↓    ↓    ↓

&#x20;    Text Shapes Tables Media

&#x20;      │    │    │    │

&#x20;      └────┴────┴────┘

&#x20;           ↓

&#x20;      OOXML / COM

&#x20;           ↓

&#x20;      Render / Validate

&#x20;           ↓

&#x20;         Save

