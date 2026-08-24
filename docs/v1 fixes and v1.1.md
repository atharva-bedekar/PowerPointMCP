We have completed a real end-to-end test of the PowerPoint MCP and identified one serious correctness bug, several minor usability/efficiency issues, and a set of improvements for v1.1.



Treat this as an implementation task in two explicit phases.



PHASE 1 — FIX BUGS AND MINOR IMPROVEMENTS FOR V1



Do not redesign the architecture in this phase. Preserve the existing working implementation wherever possible.



==================================================

1\. FIX THE ACTIVE SESSION / PRESENTATION PATH BUG

==================================================



This is the highest-priority fix.



Current failure mode:



1\. ppt\_open() creates:

&#x20;  .ppt-agent/sessions/<session\_id>/working.pptx



2\. A mutation tool such as:

&#x20;  - ppt\_modify\_shape

&#x20;  - ppt\_modify\_text

&#x20;  - ppt\_delete\_shape



&#x20;  is called with an explicit presentation\_path pointing to the original source presentation.



3\. The mutation is applied to the original presentation instead of the session working copy.



4\. ppt\_save or ppt\_save\_as then reads the unchanged working.pptx and writes it back to the destination.



5\. The already-applied edits are silently lost.



This is unacceptable because rendering can appear correct before save, while the final saved PPTX is reverted.



IMPLEMENTATION REQUIREMENT:



Once ppt\_open() establishes an active session, all mutation operations must target the active session working copy.



Do not rely only on the Antigravity skill to enforce this. Enforce it in the MCP implementation.



Preferred behavior:



\- If an active session exists:

&#x20; - mutation tools always operate on session.working\_path

&#x20; - an explicit presentation\_path must not override the active session

\- If an explicit presentation\_path conflicts with the active session:

&#x20; - do not silently modify another file

&#x20; - either ignore the explicit path in favor of the active session OR return a clear validation error

\- Do not allow two mutable targets inside one active session.



Use a single canonical target-resolution mechanism for every tool.



Create something conceptually equivalent to:



resolve\_active\_target(

&#x20;   presentation\_path=None,

&#x20;   require\_session=False,

&#x20;   mutation=False

)



All mutation tools must use the same target-resolution logic.



Do not duplicate path-selection logic across individual tools.



Also make the state explicit in the response where useful, for example:



{

&#x20; "session\_id": "...",

&#x20; "target": "working",

&#x20; "path": ".../.ppt-agent/sessions/.../working.pptx"

}



This is especially useful for debugging.



==================================================

2\. ADD SESSION STATE INTEGRITY CHECKS

==================================================



Strengthen save safety.



Before ppt\_save or ppt\_save\_as:



\- Verify the session working file exists.

\- Verify the active session is valid.

\- Verify that mutations were applied to the working copy.

\- Track whether the working copy has been modified.

\- Track a working-file hash or equivalent state marker when practical.



If the server detects an impossible or suspicious state, fail clearly instead of silently overwriting files.



For example, detect cases where:



\- a session exists

\- mutations were expected

\- working.pptx has not changed

\- but another presentation file appears to have been modified



Do not attempt magical recovery. Fail safely and report the state.



==================================================

3\. STRENGTHEN THE ANTIGRAVITY SKILL

==================================================



Update:



.agents/skills/powerpoint-editor/SKILL.md



Add an explicit immutable rule near the beginning:



"After ppt\_open establishes an active session, never pass presentation\_path to mutation tools. All mutations must operate on the active session working copy. presentation\_path is for opening/initializing a session and non-session operations only."



Also update the lifecycle language to make the intended flow unambiguous:



ppt\_open

→ active session

→ inspect

→ mutate working copy

→ validate

→ render

→ verify

→ save/save\_as



Do not make the skill depend on the agent remembering subtle implementation details. The MCP itself must enforce the rule.



==================================================

4\. REDUCE DEFAULT INSPECTION OUTPUT

==================================================



The current ppt\_inspect\_slide and ppt\_validate\_slide can produce very large JSON outputs on complex slides.



Keep all detailed information available, but make the default output concise and agent-friendly.



The default ppt\_inspect\_slide result should prioritize:



\- slide dimensions

\- shape count

\- important semantic roles

\- shape IDs

\- x/y/width/height

\- brief text summary

\- important style information

\- warnings



Do not dump unnecessary raw XML or enormous per-run data by default.



Preserve a detailed/deep inspection mode.



For example, support a parameter conceptually equivalent to:



detail="summary"

detail="full"



Default to summary.



ppt\_inspect\_shape can remain the primary way to retrieve deep information about a single object.



Do not remove existing information from the implementation. Make the normal result smaller.



==================================================

5\. IMPROVE VALIDATION OUTPUT

==================================================



Keep all current validation rules, but make the normal result concise and structured.



Default output should contain:



\- valid / invalid

\- counts by issue type

\- critical issues

\- shape IDs involved

\- concise descriptions



For example:



{

&#x20; "valid": false,

&#x20; "summary": {

&#x20;   "overlaps": 2,

&#x20;   "boundary\_violations": 0,

&#x20;   "text\_overflow": 1,

&#x20;   "tiny\_fonts": 0

&#x20; },

&#x20; "issues": \[...]

}



Do not return thousands of lines unless detailed mode was explicitly requested.



==================================================

6\. SAVE SAFETY

==================================================



Preserve the existing backup behavior.



Keep automatic timestamped backups.



Make save\_as the preferred non-destructive workflow.



Do not change existing behavior unnecessarily, but ensure the skill and documentation make clear:



\- ppt\_save may write back to the original

\- ppt\_save\_as creates a separate output file

\- originals should not be overwritten unless explicitly requested



If a default output directory already exists in configuration, keep it.



Do not invent a new output architecture in v1.



==================================================

7\. ADD BASIC TESTS FOR THE BUG

==================================================



Create a regression test that reproduces the exact failure:



1\. Open source presentation.

2\. Create session.

3\. Modify using an explicit original presentation\_path.

4\. Save.

5\. Verify that the saved result contains the modification.



Then add the stricter test:



1\. Open source presentation.

2\. Create session.

3\. Perform mutation while omitting presentation\_path.

4\. Save.

5\. Verify modification exists in output.



Then test:



\- explicit path conflicting with active session

\- save\_as

\- save

\- revert

\- backup creation



The session bug must not be considered fixed until these tests pass.



==================================================

8\. DO NOT MAKE THESE V1 CHANGES

==================================================



Do not implement the following in Phase 1 unless necessary to fix an existing bug:



\- major MCP redesign

\- new UI

\- remote server

\- LLM inside the MCP

\- PowerPoint generation framework

\- high-level diagram engine

\- large new tool family

\- broad refactoring unrelated to the identified issues



The goal is a stable and safe v1.



At the end of Phase 1:



\- run the full test suite

\- run an end-to-end PowerPoint edit

\- verify rendering

\- verify validation

\- verify save

\- verify the saved file contains the edits

\- verify the original is preserved when using save\_as



Provide a concise Phase 1 report including:

\- files changed

\- bugs fixed

\- tests added

\- test results

\- any remaining known limitations





==================================================

PHASE 2 — DESIGN AND IMPLEMENT V1.1 FEATURES

==================================================



After Phase 1 is complete and verified, implement the following v1.1 improvements.



The purpose of v1.1 is not to change the underlying deterministic PowerPoint editing philosophy.



The purpose is to make the MCP substantially more efficient for an LLM agent performing multi-object edits and slide construction.



==================================================

V1.1 FEATURE 1 — BATCH SHAPE MUTATION

==================================================



Add a batch mutation capability.



Preferred tool:



ppt\_batch\_modify\_shapes



It should allow the agent to modify multiple shapes in one MCP call.



Each operation should support the same safe properties already supported by ppt\_modify\_shape, including where appropriate:



\- x

\- y

\- width

\- height

\- rotation

\- z\_order

\- dx

\- dy

\- dwidth

\- dheight

\- drotation

\- alignment/distribution behavior



Example conceptual input:



{

&#x20; "slide\_number": 3,

&#x20; "operations": \[

&#x20;   {

&#x20;     "shape\_id": 13,

&#x20;     "changes": {

&#x20;       "height": 5.4

&#x20;     }

&#x20;   },

&#x20;   {

&#x20;     "shape\_id": 14,

&#x20;     "changes": {

&#x20;       "y": 2.15

&#x20;     }

&#x20;   },

&#x20;   {

&#x20;     "shape\_id": 15,

&#x20;     "changes": {

&#x20;       "y": 2.15

&#x20;     }

&#x20;   }

&#x20; ]

}



Requirements:



\- execute deterministically

\- validate shape IDs before mutation

\- produce structured per-operation results

\- fail safely if an operation cannot be applied

\- keep the entire batch within the same session working copy

\- do not create partial corruption



The response should summarize what changed rather than returning a massive object dump.



==================================================

V1.1 FEATURE 2 — BASIC SHAPE CREATION

==================================================



Add tools for creating new PowerPoint objects.



At minimum:



ppt\_add\_shape

ppt\_add\_textbox

ppt\_add\_connector



Do not attempt to create every possible PowerPoint feature.



Support common objects first.



ppt\_add\_shape should support common geometric types such as:



\- rectangle

\- rounded\_rectangle

\- ellipse

\- line

\- arrow

\- chevron



Support properties such as:



\- x

\- y

\- width

\- height

\- rotation

\- fill

\- line

\- line width

\- transparency where already supported by the OOXML helpers

\- style inheritance when a reference shape is provided



ppt\_add\_textbox should support:



\- text

\- x

\- y

\- width

\- height

\- font family

\- font size

\- bold

\- italic

\- underline

\- color

\- alignment

\- margins



ppt\_add\_connector should support:



\- source shape

\- target shape

\- connector type

\- arrow end/start

\- basic line styling



Prefer deterministic PowerPoint-native objects over rasterized images.



==================================================

V1.1 FEATURE 3 — BATCH EDITING / COMBINED MUTATIONS

==================================================



Evaluate whether a general-purpose:



ppt\_batch\_edit



is preferable to multiple specialized batch tools.



Do not automatically add it if it duplicates ppt\_batch\_modify\_shapes unnecessarily.



The design goal is:



\- one MCP call can perform a logically related set of deterministic edits

\- the LLM should not need 10 round trips for 10 independent shape changes

\- the API must remain understandable to the model

\- operations must be validated before mutation when practical



Favor a small, clean tool surface over a proliferation of batch tools.



==================================================

V1.1 FEATURE 4 — HIGH-LEVEL FLOW DIAGRAM CREATION

==================================================



Add a composite capability:



ppt\_create\_flow\_diagram



This is specifically intended for prompts such as:



"Create a 10-step flow diagram with these steps."



The MCP should internally perform the low-level PowerPoint operations.



The agent should not need to individually create every rectangle and connector.



Input should conceptually support:



\- slide number

\- list of steps

\- layout direction

\- number of rows/columns where applicable

\- spacing

\- node dimensions

\- connector style

\- reference/style source where practical



Example:



{

&#x20; "slide\_number": 5,

&#x20; "steps": \[

&#x20;   "Request",

&#x20;   "Validate",

&#x20;   "Analyze",

&#x20;   "Plan",

&#x20;   "Approve",

&#x20;   "Configure",

&#x20;   "Deploy",

&#x20;   "Test",

&#x20;   "Monitor",

&#x20;   "Complete"

&#x20; ],

&#x20; "direction": "horizontal"

}



The implementation should:



1\. calculate geometry

2\. create nodes

3\. create connectors

4\. align and distribute them

5\. apply consistent styling

6\. return the created shape IDs and bounding region



Do not build an entire diagramming framework in v1.1.



Implement a reliable, useful first version.



==================================================

V1.1 FEATURE 5 — STYLE INHERITANCE / REFERENCE OBJECTS

==================================================



Improve creation/editing workflows by allowing a newly created object to inherit styling from an existing object.



For example:



ppt\_add\_shape(..., style\_source\_shape\_id=14)



The new object should inherit relevant properties such as:



\- fill

\- line

\- line width

\- font

\- font size

\- text color

\- paragraph alignment

\- margins



Only inherit properties that are supported safely.



Do not modify the source object.



This is especially important for creating new content that visually belongs to an existing deck.



==================================================

V1.1 FEATURE 6 — BETTER STRUCTURAL COMPARISON

==================================================



Improve ppt\_compare\_slides so that structural comparison is the primary signal, not raw pixel similarity.



Comparison should emphasize:



\- semantic roles

\- shape correspondence

\- geometry deltas

\- typography

\- relative alignment

\- dimensions

\- spacing

\- missing/extra objects



Return a concise summary such as:



{

&#x20; "similarity": 0.92,

&#x20; "matched\_shapes": 14,

&#x20; "missing\_shapes": 1,

&#x20; "extra\_shapes": 0,

&#x20; "geometry\_differences": 3,

&#x20; "typography\_differences": 1

}



Retain detailed comparison information behind a detailed mode.



Do not treat pixel similarity as the sole correctness metric.



==================================================

V1.1 FEATURE 7 — RENDER/VERIFY EFFICIENCY

==================================================



Preserve the existing visual verification loop.



Optimize it so that:



\- only affected slides are rendered

\- full presentation rendering is not performed unless requested

\- validation results are concise

\- repeated identical rendering is avoided when practical

\- the agent can clearly determine whether another correction is required



Do not remove the current PowerPoint COM high-fidelity renderer or LibreOffice fallback.



==================================================

V1.1 FEATURE 8 — AGENT TOOL-CALL OPTIMIZATION

==================================================



Update SKILL.md to teach the agent to use the new abstractions.



The skill should explicitly prefer:



\- inspect once, then reuse inspection data

\- deep inspection only when necessary

\- batch mutations for multiple independent edits

\- composite operations for structured tasks

\- render only affected slides

\- validate before and/or after visual verification where appropriate

\- save only after verification



The guiding principle should be:



"Use the highest-level deterministic tool that safely expresses the requested operation."



Examples:



User:

"Move these eight boxes down."



Prefer:

ppt\_batch\_modify\_shapes



Not:

ppt\_modify\_shape × 8



User:

"Create a 10-step flow."



Prefer:

ppt\_create\_flow\_diagram



Not:

ppt\_add\_shape × 10 + ppt\_add\_connector × 9



User:

"Change the title."



Prefer:

ppt\_modify\_text



Not:

delete + recreate



==================================================

V1.1 FEATURE 9 — AGENT-FRIENDLY TOOL RESPONSES

==================================================



Review all major MCP responses.



Every tool should return concise structured information sufficient for the next agent decision.



Avoid:



\- giant raw XML

\- redundant shape dumps

\- repeated data already known

\- huge validation payloads

\- unnecessary filesystem details



Responses should summarize:



\- what happened

\- what changed

\- relevant IDs

\- warnings

\- whether the operation succeeded

\- what the agent should do next if additional work is required



Do not sacrifice access to detailed information. Provide detail mode when necessary.



==================================================

V1.1 FEATURE 10 — END-TO-END PERFORMANCE BENCHMARK

==================================================



Create a benchmark for agent efficiency.



Use at least these scenarios:



SCENARIO A:

"Move the title 0.2 inches left."



SCENARIO B:

"Make these eight boxes evenly distributed."



SCENARIO C:

"Remove two boxes and reposition the remaining elements."



SCENARIO D:

"Create a 10-step flow diagram."



SCENARIO E:

"Make slide 7 match slide 3 while preserving its content."



For each scenario measure:



\- MCP tool calls

\- rendering calls

\- validation calls

\- total agent turns if measurable

\- size of MCP responses

\- correction iterations

\- final correctness



The purpose is to determine whether v1.1 actually reduces unnecessary tool calls.



Do not optimize for a specific arbitrary call-count target at the expense of correctness.



==================================================

V1.1 TEST REQUIREMENTS

==================================================



Add tests for:



\- batch shape mutation

\- shape creation

\- textbox creation

\- connector creation

\- flow diagram creation

\- style inheritance

\- structural slide comparison

\- concise inspection responses

\- concise validation responses

\- session safety across all new tools

\- rendering of newly created objects

\- save/save\_as after batch operations

\- rollback/revert after failed batch operation



Also run the original real-world edit scenario:



"On Slide 3. remove the Orchestration box, and its adjacent box. Extend the Client Configuration box to the bottom of the slide, and also add content and extend the box adjacent to the Client Configuration box"



Compare the result against the previous run.



The v1.1 implementation should require substantially fewer mutation round trips for this type of multi-object edit while preserving the same visual result.



==================================================

FINAL DELIVERABLE

==================================================



After Phase 1 and Phase 2, provide a report with:



1\. Phase 1 bugs fixed.

2\. Phase 1 tests added and results.

3\. Phase 2 features implemented.

4\. New MCP tools.

5\. Changes to SKILL.md.

6\. Tool-call efficiency improvements.

7\. Before/after benchmark results.

8\. Any remaining limitations.

9\. Exact commands to test the MCP.

10\. Confirmation that the end-to-end PowerPoint editing scenarios were actually executed.



IMPORTANT IMPLEMENTATION PRINCIPLES



\- Do not rewrite working components unnecessarily.

\- Preserve deterministic PowerPoint manipulation.

\- Preserve the working-copy/session architecture.

\- Preserve high-fidelity PowerPoint COM rendering.

\- Preserve LibreOffice fallback.

\- Preserve automatic backups.

\- Prefer small, composable, agent-friendly MCP APIs.

\- Make the MCP enforce safety instead of relying solely on model instructions.

\- Optimize for both correctness and agent efficiency.

\- Do not optimize merely by reducing tool count if it makes the resulting presentation less reliable.

\- Do not consider a feature complete until it has been implemented and tested with an actual PPTX.

