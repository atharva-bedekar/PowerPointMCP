# Core Engine Technical Specification Handoff (R1 & R2)

**Agent ID**: `spec_miner_core_1`  
**Milestone**: R1 (Deterministic PPTX Inspection & Geometry Engine) & R2 (Rendering & Visual Verification Pipeline)  
**Date**: 2026-08-21T05:58:00Z  

---

## 1. Observation

### 1.1 Environment & Platform Capabilities
- **OS**: Windows 11 (AMD64)
- **Runtime**: Python 3.14.3 managed via `uv 0.10.11`.
- **Microsoft PowerPoint**: Installed at `C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE` (verified via registry `HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe`).
- **PowerPoint COM Automation**: Verified functional using `win32com.client.DispatchEx('PowerPoint.Application')` and `pythoncom.CoInitialize()`. Slide export via `presentation.Slides[i].Export(path, 'PNG', width, height)` executes headlessly with `WithWindow=False` in ~300ms.
- **Python Libraries**: `python-pptx` (0.6.23+), `Pillow` (12.3.0), `numpy` (2.4.3), `pywin32` / `comtypes`, and `lxml` verified installable and fully compatible.
- **LibreOffice**: `soffice.exe` not currently in standard `App Paths`, establishing PowerPoint COM as the primary renderer on this machine, with LibreOffice headless CLI as the cross-platform fallback.

### 1.2 Underlying PPTX & OpenXML Realities
- **EMU Unit Standard**: ECMA-376 specifies English Metric Units (EMU). $1\text{ inch} = 914,400\text{ EMU}$, $1\text{ pt} = 12,700\text{ EMU}$, $1\text{ cm} = 360,000\text{ EMU}$. Slide coordinates and shape dimensions in python-pptx are stored as integer EMUs.
- **Z-Order in OpenXML**: Visual z-order is determined strictly by the child element index inside the slide's shape tree `<p:spTree>`. The first shapes in the XML render at the back; later shapes render on top. Reordering requires manipulating `<p:spTree>` child nodes.
- **Shape Duplication & Relationships**: Replicating shapes (especially pictures and complex shapes) requires deep-copying `<p:sp>` / `<p:pic>` XML elements, generating unique `<p:cNvPr id="...">` identifiers, and replicating underlying OpenXML part relationships (`r:embed` / `r:id`) in `.rels` to prevent corrupted presentations.
- **Text Frame Run Hierarchy**: PowerPoint text structure is `TextFrame -> Paragraph -> Run`. Replacing text at the paragraph level destroys all run-level formatting unless styles are explicitly captured and reapplied or edits are performed directly on specific runs.

---

## 2. Logic Chain

1. **Precision & Idempotency**: Repeated floating-point conversions between inches and EMUs introduce cumulative drift. Therefore, all internal geometry computations (bounding boxes, intersections, alignments, distributions) must be executed using integer EMUs, and converted to floating-point inches (rounded to 4 decimal places) only when formatting outputs for the agent/user.
2. **Conservative Semantic Role Inference**: Agents require high-level semantic hints (`title`, `subtitle`, `body`, `image`, `diagram`, `footer`) to understand slides quickly. Inference must use a multi-tiered hierarchy: (1) native PowerPoint placeholder type (`PP_PLACEHOLDER`), (2) shape type & element composition, (3) spatial positioning on slide ($y$-coordinate relative to slide height), and (4) typographical cues (font size, weight).
3. **Run-Level Style Preservation**: To avoid destroying typography when editing text, the engine must support both whole-textbox text replacement (propagating base run style) and targeted substring/run modifications.
4. **Deterministic Visual Diffing**: To provide immediate verification feedback without external cloud APIs, visual comparison must execute locally using numpy pixel matrix subtraction, threshold masking, connected-region clustering, and Pillow highlight compositing.
5. **Robust COM Lifecycle Management**: COM automation on Windows is vulnerable to hanging `POWERPNT.EXE` orphan processes if exceptions occur. The renderer must strictly follow `CoInitialize()` -> `DispatchEx` -> `try...finally` -> `presentation.Close()` -> `app.Quit()` -> `CoUninitialize()` -> `gc.collect()`.

---

## 3. Caveats & Assumptions

- **PowerPoint COM Concurrency**: PowerPoint COM automation is Single-Threaded Apartment (STA). Concurrent calls to PowerPoint COM from multiple threads must be serialized or initialized with appropriate apartment threading models.
- **LibreOffice Rendering Fidelity**: LibreOffice headless rendering may exhibit minor font rendering and layout differences compared to native Microsoft PowerPoint COM. The renderer metadata must explicitly state which renderer was utilized.
- **Group Shape Coordinates**: Shapes inside a group shape (`<p:grpSp>`) have coordinates relative to the group's coordinate system (`<a:chOff>` / `<a:chExt>`), not the slide origin. Inspection and editing must account for group hierarchies.
- **Zero-Area Shapes & Connectors**: Lines and connectors have either 0 width or 0 height in their bounding boxes, which require boundary epsilon handling during intersection calculations.

---

## 4. Technical Specifications

```
================================================================================
SECTION A: DATA MODELS & SCHEMA DEFINITIONS
================================================================================
```

### A.1 Enums

```python
from enum import Enum

class SemanticRole(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    BODY = "body"
    IMAGE = "image"
    DIAGRAM = "diagram"
    TABLE = "table"
    CHART = "chart"
    FOOTER = "footer"
    UNKNOWN = "unknown"

class ShapeType(str, Enum):
    AUTO_SHAPE = "auto_shape"
    TEXT_BOX = "text_box"
    PICTURE = "picture"
    GROUP = "group"
    TABLE = "table"
    CHART = "chart"
    CONNECTOR = "connector"
    MEDIA = "media"
    UNKNOWN = "unknown"

class AlignmentType(str, Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"

class DistributionMode(str, Enum):
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"

class SpacingMode(str, Enum):
    EQUAL_GAPS = "equal_gaps"
    EQUAL_CENTERS = "equal_centers"
```

### A.2 Core Data Models

```python
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

@dataclass
class BoundingBox:
    left_emu: int
    top_emu: int
    width_emu: int
    height_emu: int

    @property
    def right_emu(self) -> int:
        return self.left_emu + self.width_emu

    @property
    def bottom_emu(self) -> int:
        return self.top_emu + self.height_emu

    @property
    def center_x_emu(self) -> int:
        return self.left_emu + self.width_emu // 2

    @property
    def center_y_emu(self) -> int:
        return self.top_emu + self.height_emu // 2

    @property
    def left_inches(self) -> float:
        return round(self.left_emu / 914400.0, 4)

    @property
    def top_inches(self) -> float:
        return round(self.top_emu / 914400.0, 4)

    @property
    def width_inches(self) -> float:
        return round(self.width_emu / 914400.0, 4)

    @property
    def height_inches(self) -> float:
        return round(self.height_emu / 914400.0, 4)

    @property
    def right_inches(self) -> float:
        return round(self.right_emu / 914400.0, 4)

    @property
    def bottom_inches(self) -> float:
        return round(self.bottom_emu / 914400.0, 4)

    def to_dict(self) -> Dict[str, float]:
        return {
            "x": self.left_inches,
            "y": self.top_inches,
            "width": self.width_inches,
            "height": self.height_inches,
            "right": self.right_inches,
            "bottom": self.bottom_inches,
        }

@dataclass
class TextStyle:
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color_rgb: Optional[str] = None  # Hex format: "#RRGGBB" or "RRGGBB"
    alignment: Optional[str] = None   # "left", "center", "right", "justify"
    line_spacing_pt: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

@dataclass
class TextRunModel:
    text: str
    style: TextStyle

@dataclass
class ParagraphModel:
    text: str
    runs: List[TextRunModel] = field(default_factory=list)
    alignment: Optional[str] = None
    level: int = 0
    line_spacing_pt: Optional[float] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

@dataclass
class TextFrameModel:
    text: str
    paragraphs: List[ParagraphModel] = field(default_factory=list)
    word_wrap: bool = True
    margin_left_inches: float = 0.1
    margin_right_inches: float = 0.1
    margin_top_inches: float = 0.05
    margin_bottom_inches: float = 0.05
    vertical_anchor: Optional[str] = None

@dataclass
class ShapeModel:
    shape_id: int
    name: str
    shape_type: ShapeType
    semantic_role: SemanticRole
    bounds: BoundingBox
    rotation: float = 0.0
    z_order: int = 0
    text_frame: Optional[TextFrameModel] = None
    fill_color: Optional[str] = None
    fill_type: Optional[str] = None  # "solid", "gradient", "none", "pattern"
    line_color: Optional[str] = None
    line_width_pt: Optional[float] = None
    group_id: Optional[int] = None
    image_metadata: Optional[Dict[str, Any]] = None
    table_metadata: Optional[Dict[str, Any]] = None
    chart_metadata: Optional[Dict[str, Any]] = None

    def to_dict(self) -> Dict[str, Any]:
        res = {
            "shape_id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type.value,
            "role": self.semantic_role.value,
            "x": self.bounds.left_inches,
            "y": self.bounds.top_inches,
            "width": self.bounds.width_inches,
            "height": self.bounds.height_inches,
            "rotation": round(self.rotation, 2),
            "z_order": self.z_order,
            "fill": self.fill_color,
            "line": self.line_color,
        }
        if self.text_frame:
            res["text"] = self.text_frame.text
            # Extract dominant text style from first run
            if self.text_frame.paragraphs and self.text_frame.paragraphs[0].runs:
                first_style = self.text_frame.paragraphs[0].runs[0].style
                res["font_family"] = first_style.font_name
                res["font_size"] = first_style.font_size_pt
                res["bold"] = first_style.bold
                res["italic"] = first_style.italic
                res["color"] = first_style.color_rgb
                res["alignment"] = self.text_frame.paragraphs[0].alignment
        if self.image_metadata:
            res["image_metadata"] = self.image_metadata
        if self.table_metadata:
            res["table_metadata"] = self.table_metadata
        if self.chart_metadata:
            res["chart_metadata"] = self.chart_metadata
        return res

@dataclass
class SlideModel:
    slide_number: int
    slide_id: int
    layout_name: str
    title: Optional[str] = None
    shapes: List[ShapeModel] = field(default_factory=list)
    has_notes: bool = False
    notes_text: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "slide_id": self.slide_id,
            "layout_name": self.layout_name,
            "title": self.title,
            "shape_count": len(self.shapes),
            "shapes": [s.to_dict() for s in self.shapes],
            "has_notes": self.has_notes,
        }

@dataclass
class PresentationMetadata:
    title: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None
    revision: Optional[int] = None

@dataclass
class PresentationModel:
    presentation_path: str
    slide_count: int
    width_inches: float
    height_inches: float
    width_emu: int
    height_emu: int
    slide_titles: List[Dict[str, Any]]
    available_layouts: List[str]
    metadata: PresentationMetadata
    theme_name: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "presentation_path": self.presentation_path,
            "slide_count": self.slide_count,
            "dimensions": {
                "width_inches": self.width_inches,
                "height_inches": self.height_inches,
                "width_emu": self.width_emu,
                "height_emu": self.height_emu,
            },
            "slide_titles": self.slide_titles,
            "available_layouts": self.available_layouts,
            "metadata": {
                "title": self.metadata.title,
                "author": self.metadata.author,
                "created": self.metadata.created,
                "modified": self.metadata.modified,
            }
        }
```

---

```
================================================================================
SECTION B: UNITS & PRECISION ENGINE
================================================================================
```

### B.1 Mathematical Constants & Conversion Functions

```python
# Exact conversion factors per ECMA-376 Standard
EMU_PER_INCH = 914400
EMU_PER_POINT = 12700
EMU_PER_CM = 360000
POINTS_PER_INCH = 72

def inches_to_emu(inches: float) -> int:
    """Convert float inches to exact integer EMUs."""
    return int(round(inches * EMU_PER_INCH))

def emu_to_inches(emu: int, precision: int = 4) -> float:
    """Convert integer EMUs to float inches with rounding precision."""
    return round(float(emu) / EMU_PER_INCH, precision)

def pt_to_emu(pt: float) -> int:
    """Convert points to exact integer EMUs."""
    return int(round(pt * EMU_PER_POINT))

def emu_to_pt(emu: int, precision: int = 2) -> float:
    """Convert EMUs to points."""
    return round(float(emu) / EMU_PER_POINT, precision)

def apply_delta_inches(current_emu: int, delta_inches: float) -> int:
    """Apply a relative shift in inches to an existing EMU value without cumulative rounding error."""
    return current_emu + inches_to_emu(delta_inches)
```

---

```
================================================================================
SECTION C: PPTX INSPECTOR & SEMANTIC ROLE INFERENCE
================================================================================
```

### C.1 Semantic Role Inference Rules

The inspector applies a deterministic 5-stage rule cascade to assign a `SemanticRole` to every shape:

```
Stage 1: Placeholder Examination
  IF shape.is_placeholder:
    type == PP_PLACEHOLDER.TITLE or CENTER_TITLE -> SemanticRole.TITLE
    type == PP_PLACEHOLDER.SUBTITLE -> SemanticRole.SUBTITLE
    type == PP_PLACEHOLDER.BODY -> SemanticRole.BODY
    type in (FOOTER, SLIDE_NUMBER, DATE) -> SemanticRole.FOOTER
    type in (PICTURE, BITMAP) -> SemanticRole.IMAGE
    type == TABLE -> SemanticRole.TABLE
    type == CHART -> SemanticRole.CHART

Stage 2: Non-Placeholder Structural Types
  IF shape.shape_type == PICTURE -> SemanticRole.IMAGE
  IF shape.has_table -> SemanticRole.TABLE
  IF shape.has_chart -> SemanticRole.CHART
  IF shape.shape_type == GROUP:
     IF group contains vector shapes/connectors -> SemanticRole.DIAGRAM
     ELSE -> SemanticRole.UNKNOWN

Stage 3: Spatial & Typographical Heuristics (Text Shapes)
  IF shape.has_text_frame and text is not empty:
     norm_top = shape.top_emu / slide_height_emu
     max_font_size = max font size in runs (default 18pt)
     
     # Rule 3A: Title Detection
     IF (norm_top < 0.22 and max_font_size >= 24) or ("Title" in shape.name and norm_top < 0.35):
        RETURN SemanticRole.TITLE
        
     # Rule 3B: Subtitle Detection
     IF norm_top >= 0.15 and norm_top < 0.38 and 14 <= max_font_size < 24:
        RETURN SemanticRole.SUBTITLE
        
     # Rule 3C: Footer Detection
     IF norm_top >= 0.85 or ("Footer" in shape.name or "Slide Number" in shape.name):
        RETURN SemanticRole.FOOTER
        
     # Rule 3D: Body Content
     IF len(paragraphs) > 1 or norm_top >= 0.25:
        RETURN SemanticRole.BODY

Stage 4: Default Fallback
  RETURN SemanticRole.UNKNOWN
```

### C.2 Inspector Interface

```python
class PPTXInspector:
    @staticmethod
    def inspect_presentation(presentation_path: str) -> PresentationModel:
        """Inspects presentation metadata, dimensions, layouts, and slide titles."""
        ...

    @staticmethod
    def inspect_slide(presentation_path: str, slide_number: int) -> SlideModel:
        """Inspects a single slide (1-indexed) returning all shapes with coordinates in inches."""
        ...

    @staticmethod
    def inspect_shape(presentation_path: str, slide_number: int, shape_id: int) -> ShapeModel:
        """Inspects a specific shape by ID on a slide, extracting deep typography, fill, and geometry."""
        ...
```

---

```
================================================================================
SECTION D: GEOMETRY ENGINE SPECIFICATION
================================================================================
```

### D.1 Geometric Calculations & Algorithms

#### 1. Collision & Overlap Detection
- **Axis-Aligned Bounding Box (AABB) Intersection**:
  $$\text{Overlap}_X = \max(0, \min(A.\text{right}, B.\text{right}) - \max(A.\text{left}, B.\text{left}))$$
  $$\text{Overlap}_Y = \max(0, \min(A.\text{bottom}, B.\text{bottom}) - \max(A.\text{top}, B.\text{top}))$$
  $$\text{Intersecting} \iff \text{Overlap}_X > 0 \land \text{Overlap}_Y > 0$$
- **Overlap Area in Square Inches**:
  $$\text{Area} = \frac{\text{Overlap}_X \cdot \text{Overlap}_Y}{(914,400)^2}$$

#### 2. Shape Alignment
Given a list of $n$ shapes $[S_1, S_2, \dots, S_n]$ and optional reference shape $R$:
- `LEFT`: Set each $S_i.\text{left} = R.\text{left}$ (or $\min_{j} S_j.\text{left}$).
- `CENTER`: Set each $S_i.\text{left} = R.\text{center\_x} - \frac{S_i.\text{width}}{2}$ (or $\text{avg\_center\_x} - \frac{S_i.\text{width}}{2}$).
- `RIGHT`: Set each $S_i.\text{left} = R.\text{right} - S_i.\text{width}$ (or $\max_j S_j.\text{right} - S_i.\text{width}$).
- `TOP`: Set each $S_i.\text{top} = R.\text{top}$ (or $\min_j S_j.\text{top}$).
- `MIDDLE`: Set each $S_i.\text{top} = R.\text{center\_y} - \frac{S_i.\text{height}}{2}$ (or $\text{avg\_center\_y} - \frac{S_i.\text{height}}{2}$).
- `BOTTOM`: Set each $S_i.\text{top} = R.\text{bottom} - S_i.\text{height}$ (or $\max_j S_j.\text{bottom} - S_i.\text{height}$).

#### 3. Shape Distribution
- **Horizontal Equal Gaps**:
  1. Sort shapes in ascending order of `left_emu`: $[S_{(1)}, S_{(2)}, \dots, S_{(n)}]$.
  2. Total span $W_{\text{span}} = S_{(n)}.\text{right} - S_{(1)}.\text{left}$.
  3. Total shape widths $W_{\text{shapes}} = \sum_{i=1}^n S_{(i)}.\text{width}$.
  4. Exact Gap: $\text{gap\_emu} = \lfloor \frac{W_{\text{span}} - W_{\text{shapes}}}{n - 1} \rfloor$.
  5. Position for $i \in [2, n-1]$:
     $$S_{(i)}.\text{left} = S_{(i-1)}.\text{left} + S_{(i-1)}.\text{width} + \text{gap\_emu}$$
- **Vertical Equal Gaps**:
  1. Sort by `top_emu`.
  2. Total span $H_{\text{span}} = S_{(n)}.\text{bottom} - S_{(1)}.\text{top}$.
  3. Gap: $\text{gap\_emu} = \lfloor \frac{H_{\text{span}} - \sum S_{(i)}.\text{height}}{n - 1} \rfloor$.
  4. Position $S_{(i)}.\text{top} = S_{(i-1)}.\text{bottom} + \text{gap\_emu}$.

#### 4. Dimension Equalization
- `equalize_widths(shapes, mode='first'|'max'|'min'|'avg'|target_value)`:
  Sets $S_i.\text{width} = W_{\text{target}}$ in integer EMUs.
- `equalize_heights(shapes, mode='first'|'max'|'min'|'avg'|target_value)`:
  Sets $S_i.\text{height} = H_{\text{target}}$ in integer EMUs.

#### 5. Slide Boundary & Overflow Validation
- Slide bounds: $(0, 0, W_{\text{slide}}, H_{\text{slide}})$.
- Breaches:
  - Left breach: $S.\text{left} < 0 \implies \text{breach} = |S.\text{left}|$.
  - Top breach: $S.\text{top} < 0 \implies \text{breach} = |S.\text{top}|$.
  - Right breach: $S.\text{right} > W_{\text{slide}} \implies \text{breach} = S.\text{right} - W_{\text{slide}}$.
  - Bottom breach: $S.\text{bottom} > H_{\text{slide}} \implies \text{breach} = S.\text{bottom} - H_{\text{slide}}$.

---

```
================================================================================
SECTION E: EDITING OPERATIONS & RUN-LEVEL STYLE PRESERVATION
================================================================================
```

### E.1 Modify Shape Engine (`modify_shape`)

Supports atomic updates for absolute coordinates or relative offsets.
```python
def modify_shape(
    slide: Slide,
    shape_id: int,
    x: Optional[float] = None,       # Absolute x (inches)
    y: Optional[float] = None,       # Absolute y (inches)
    width: Optional[float] = None,   # Absolute width (inches)
    height: Optional[float] = None,  # Absolute height (inches)
    dx: Optional[float] = None,      # Delta x (inches)
    dy: Optional[float] = None,      # Delta y (inches)
    dwidth: Optional[float] = None,  # Delta width (inches)
    dheight: Optional[float] = None, # Delta height (inches)
    rotation: Optional[float] = None,# Rotation (degrees)
    drotation: Optional[float] = None
) -> ShapeModel:
```

### E.2 Run-Level Style Preservation Algorithm (`modify_text`)

When modifying text in a shape, we must NOT destroy existing font weights, colors, or sizes.

```python
def modify_text(
    shape: BaseShape,
    new_text: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size_pt: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color_rgb: Optional[str] = None,
    alignment: Optional[str] = None,
    line_spacing_pt: Optional[float] = None,
    space_before_pt: Optional[float] = None,
    space_after_pt: Optional[float] = None,
    margin_left_inches: Optional[float] = None,
    margin_right_inches: Optional[float] = None,
    margin_top_inches: Optional[float] = None,
    margin_bottom_inches: Optional[float] = None,
    preserve_runs: bool = True
) -> None:
```

**Algorithm Execution**:
1. **Text Replacement with Single Run**:
   If `new_text` is provided and the paragraph has a single run, directly set `run.text = new_text`. This preserves 100% of formatting properties.
2. **Text Replacement with Multiple Runs**:
   - Capture base formatting properties from the first run (`font.name`, `font.size`, `font.bold`, `font.italic`, `font.color`).
   - Clear paragraph text (`paragraph.text = new_text`).
   - Reapply the captured properties to `paragraph.runs[0]`.
3. **Format Property Overrides**:
   - Only explicitly provided formatting parameters (non-`None`) mutate the run or paragraph.
   - All unmentioned properties remain untouched.

### E.3 Shape Copying, Deletion, and Z-Order Engine

#### Copy Shape (`copy_shape`):
1. Deep-copy the shape's XML element: `new_elem = copy.deepcopy(shape._element)`.
2. Generate next available unique shape ID: `new_id = max(s.shape_id for s in slide.shapes) + 1`.
3. Update `<p:cNvPr id="...">` and `<p:cNvPr name="...">`.
4. Apply spatial offset (`x_offset_inches`, `y_offset_inches`) to `<a:off x="..." y="..." />`.
5. **Relationship Preservation**: If shape is a picture or media object with relationship ID `rId`, copy the relationship from the source slide part to the target slide part via `slide.part.relate_to(image_part, reltype)`.
6. Append `new_elem` to `slide.shapes._spTree`.

#### Delete Shape (`delete_shape`):
1. Locate target shape by ID.
2. Extract parent element: `parent = shape._element.getparent()`.
3. Remove node: `parent.remove(shape._element)`.

#### Z-Order Modification (`reorder_z_order`):
- Actions: `bring_to_front`, `send_to_back`, `bring_forward`, `send_backward`.
- Inspect `<p:spTree>` child elements. Skip non-shape header nodes (`nvGrpSpPr`, `grpSpPr`).
- Relocate `<p:sp>` / `<p:pic>` node to appropriate child index in `<p:spTree>`.

---

```
================================================================================
SECTION F: SAFE OOXML FALLBACK HELPERS
================================================================================
```

### F.1 OpenXML Namespaces & Helper Functions (`ooxml.py`)

```python
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def set_shape_gradient_fill(shape: BaseShape, start_hex: str, end_hex: str, angle_deg: float = 90.0) -> None:
    """Configures smooth two-stop gradient fill directly in OpenXML <p:spPr>."""
    ...

def set_shape_transparency(shape: BaseShape, alpha_percent: float) -> None:
    """Injects <a:alpha val="..."> into <a:solidFill>."""
    ...

def set_shape_shadow_effect(shape: BaseShape, blur_rad_pt: float, dist_pt: float, dir_deg: float) -> None:
    """Configures outer drop shadow in <a:effectLst>/<a:outerShdw>."""
    ...
```

---

```
================================================================================
SECTION G: SEMANTIC SHAPE MATCHING ALGORITHM (`match_shapes`)
================================================================================
```

### G.1 Scoring Matrix & Multi-Factor Heuristic

Given slide $A$ with shapes $[a_1, \dots, a_m]$ and slide $B$ with shapes $[b_1, \dots, b_n]$:
$$\text{Similarity}(a_i, b_j) = \sum_{k=1}^6 w_k \cdot S_k(a_i, b_j)$$

| Factor $k$ | Feature | Weight $w_k$ | Scoring Logic $S_k(a_i, b_j)$ |
|---|---|---|---|
| 1 | **Semantic Role** | $0.25$ | $1.0$ if $\text{role}_a = \text{role}_b \neq \text{unknown}$; $0.0$ if distinct known roles; $0.05$ if unknown. |
| 2 | **Text Similarity** | $0.25$ | Token Set Ratio / Levenshtein similarity on normalized lowercased text $\in [0.0, 1.0]$. $0.15$ if both empty. |
| 3 | **Relative Position** | $0.20$ | Normalized center Euclidean distance $d$: $\max(0.0, 1.0 - 2.0 \cdot d)$. |
| 4 | **Shape Type** | $0.15$ | $1.0$ if $\text{type}_a = \text{type}_b$, else $0.0$. |
| 5 | **Relative Dimensions** | $0.10$ | $1.0 - \frac{|\Delta w / W| + |\Delta h / H|}{2} \in [0.0, 1.0]$. |
| 6 | **Shape Name Match** | $0.05$ | Levenshtein similarity of shape names (e.g. "Title 1" vs "Title 2"). |

### G.2 Matching Solver
- Constructs $M \in \mathbb{R}^{m \times n}$.
- Applies greedy bipartite matching with threshold filtering ($\text{min\_confidence} = 0.40$).
- Returns ordered list of matches with full confidence scores and human-readable reasoning strings.

---

```
================================================================================
SECTION H: RENDERING PIPELINE & COM LIFECYCLE
================================================================================
```

### H.1 Renderer Abstract Base Class

```python
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Dict, Any

class Renderer(ABC):
    @abstractmethod
    def is_available(self) -> bool:
        """Check if renderer dependencies and executables exist on host."""
        ...

    @abstractmethod
    def render_slide(self, presentation_path: Path, slide_number: int, output_path: Path, dpi: int = 150) -> Path:
        """Render a single 1-indexed slide to PNG."""
        ...

    @abstractmethod
    def render_presentation(self, presentation_path: Path, output_dir: Path, dpi: int = 150) -> List[Path]:
        """Render all presentation slides to PNG files."""
        ...

    @abstractmethod
    def get_renderer_info(self) -> Dict[str, Any]:
        """Return renderer type, version, and status."""
        ...
```

### H.2 PowerPoint COM Automation (`PowerPointRenderer`)

```python
class PowerPointRenderer(Renderer):
    def render_slide(self, presentation_path: Path, slide_number: int, output_path: Path, dpi: int = 150) -> Path:
        import pythoncom
        import win32com.client
        
        abs_in = str(presentation_path.resolve())
        abs_out = str(output_path.resolve())
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Determine pixel dimensions based on DPI (16:9 10x5.625 in @ 150 DPI = 1500x844 px)
        width_px = int(10.0 * dpi)
        height_px = int(5.625 * dpi)
        
        pythoncom.CoInitialize()
        app = None
        pres = None
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            # Open presentation invisibly
            pres = app.Presentations.Open(abs_in, ReadOnly=True, Untitled=False, WithWindow=False)
            slide = pres.Slides(slide_number)
            slide.Export(abs_out, "PNG", width_px, height_px)
            return output_path
        finally:
            if pres:
                try:
                    pres.Close()
                except Exception:
                    pass
            if app:
                try:
                    app.Quit()
                except Exception:
                    pass
            del pres
            del app
            pythoncom.CoUninitialize()
            import gc
            gc.collect()
```

### H.3 LibreOffice Headless Fallback (`LibreOfficeRenderer`)

- Command: `soffice --headless --convert-to pdf --outdir <temp_dir> <presentation_path>`
- Converts generated PDF to PNG slides using `pymupdf` (`fitz`) or Pillow at target DPI.
- Timeout protection: 30-second subprocess deadline to prevent frozen soffice instances.

---

```
================================================================================
SECTION I: VISUAL VERIFICATION & IMAGE DIFFING
================================================================================
```

### I.1 Image Comparison & Diffing Algorithm (`image_diff.py`)

1. **Load & Align**:
   Load before/after images using `PIL.Image.open()`. Ensure RGB color mode and identical pixel dimensions $(W, H)$.
2. **Numpy Pixel Subtraction**:
   $$\mathbf{D}(x, y) = \max_{c \in \{R, G, B\}} |I_{\text{after}}(x, y, c) - I_{\text{before}}(x, y, c)|$$
3. **Threshold Mask**:
   $$\mathbf{M}(x, y) = \mathbf{D}(x, y) > 15$$
4. **Metrics**:
   - $\text{Total Pixels} = W \cdot H$
   - $\text{Changed Pixels} = \sum_{x, y} \mathbf{M}(x, y)$
   - $\text{Similarity Percentage} = 100.0 \times \left(1.0 - \frac{\text{Changed Pixels}}{\text{Total Pixels}}\right)$
5. **Clustered Changed Bounding Boxes**:
   - Group changed mask pixels using a grid block kernel ($32 \times 32$ pixels).
   - Find connected components across active blocks using BFS.
   - For each component, compute exact minimum bounding box $[x_{\min}, y_{\min}, x_{\max}, y_{\max}]$.
6. **Visual Diff Artifact Generation**:
   - Background: Grayscale conversion of original image $I_{\text{before}}$ blended with white at 30% alpha.
   - Overlay: Masked changed pixels painted in bright magenta `RGBA(255, 0, 128, 255)`.
   - Outlines: Red rectangular borders `(255, 0, 0, 255)` drawn around all detected changed bounding boxes.
   - Save artifact to `.ppt-agent/sessions/<session_id>/diffs/diff_slide_X.png`.

---

```
================================================================================
SECTION J: MCP TOOL SCHEMAS & ERROR HANDLING
================================================================================
```

### J.1 Core MCP Tools for R1 & R2

#### 1. `ppt_inspect_presentation`
- **Arguments**: `presentation_path: str`
- **Returns**: Structured presentation metadata, dimensions in inches/EMUs, slide count, layout list, slide titles.

#### 2. `ppt_inspect_slide`
- **Arguments**: `presentation_path: str`, `slide_number: int` (1-indexed)
- **Returns**: Full slide model with every shape's ID, name, role, coordinates in inches, typography, fill, line, and z-order.

#### 3. `ppt_inspect_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`
- **Returns**: Detailed shape model including paragraph/run breakdown, geometry, and OpenXML attributes.

#### 4. `ppt_modify_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `x: float = None`, `y: float = None`, `width: float = None`, `height: float = None`, `dx: float = None`, `dy: float = None`, `dwidth: float = None`, `dheight: float = None`, `rotation: float = None`
- **Returns**: Updated shape model.

#### 5. `ppt_modify_text`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `text: str = None`, `font_family: str = None`, `font_size: float = None`, `bold: bool = None`, `italic: bool = None`, `underline: bool = None`, `color: str = None`, `alignment: str = None`, `line_spacing: float = None`, `space_before: float = None`, `space_after: float = None`
- **Returns**: Updated text frame model.

#### 6. `ppt_copy_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `target_slide_number: int = None`, `offset_x: float = 0.5`, `offset_y: float = 0.5`
- **Returns**: Model of the newly created shape with new `shape_id`.

#### 7. `ppt_move_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `x: float = None`, `y: float = None`, `dx: float = None`, `dy: float = None`
- **Returns**: Updated shape model.

#### 8. `ppt_resize_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `width: float = None`, `height: float = None`, `dwidth: float = None`, `dheight: float = None`
- **Returns**: Updated shape model.

#### 9. `ppt_delete_shape`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`
- **Returns**: `{"success": True, "deleted_shape_id": int}`.

#### 10. `ppt_modify_ooxml`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `shape_id: int`, `operation: str`, `parameters: dict`
- **Returns**: Operation result.

#### 11. `ppt_render_slide`
- **Arguments**: `presentation_path: str`, `slide_number: int`, `output_path: str = None`, `dpi: int = 150`
- **Returns**: `{"image_path": str, "renderer": str, "width_px": int, "height_px": int}`.

#### 12. `ppt_render_presentation`
- **Arguments**: `presentation_path: str`, `output_dir: str = None`, `dpi: int = 150`
- **Returns**: `{"rendered_slides": List[str], "renderer": str, "slide_count": int}`.

#### 13. `ppt_compare_slides`
- **Arguments**: `presentation_path: str`, `slide_a: int`, `slide_b: int`
- **Returns**: Geometric, typographical, and semantic shape matching comparison report between slide A and slide B.

#### 14. `ppt_visual_diff`
- **Arguments**: `before_image: str`, `after_image: str`, `output_diff_path: str = None`
- **Returns**: `{"similarity_percentage": float, "changed_pixels": int, "total_pixels": int, "changed_regions": List[dict], "diff_image_path": str}`.

### J.2 Structured Error Model

When an operation encounters an error, it returns a structured recovery object:

```python
@dataclass
class MCPErrorResponse:
    success: bool = False
    error_type: str  # e.g., "ShapeNotFound", "SlideIndexOutOfBounds", "InvalidCoordinate"
    message: str
    context: Dict[str, Any] = field(default_factory=dict)
```

Example JSON response:
```json
{
  "success": false,
  "error_type": "ShapeNotFound",
  "message": "Shape ID 42 does not exist on slide 2.",
  "context": {
    "requested_shape_id": 42,
    "slide_number": 2,
    "available_shapes": [
      {"shape_id": 2, "name": "Title 1", "role": "title"},
      {"shape_id": 3, "name": "Content Placeholder 2", "role": "body"}
    ]
  }
}
```

---

## 5. Verification Method

To verify implementations conforming to this specification:
1. **Unit Testing**: Run `pytest tests/test_inspection.py tests/test_geometry.py tests/test_editing.py tests/test_rendering.py`.
2. **Inspection Verification**:
   - Open programmatic synthetic presentation.
   - Assert `inspect_presentation` returns correct dimensions ($10.0 \times 5.625$ in for 16:9 or $10.0 \times 7.5$ in for 4:3).
   - Assert title shapes are assigned `SemanticRole.TITLE`.
3. **Geometry Verification**:
   - Move shape by $+0.2$ inches. Assert `shape.left_emu` increases by exactly $182,880\text{ EMU}$.
   - Horizontal distribute 3 shapes. Assert gap distance between shape 1-2 equals shape 2-3 within 1 EMU.
4. **Text Run Style Preservation**:
   - In a multi-run paragraph (e.g. bold header + normal text), update text and verify bold and color formatting attributes are preserved.
5. **Rendering & Diffing Verification**:
   - Render slide 1 via COM automation -> assert PNG output file exists and is $> 0$ bytes.
   - Move title on slide 1 and render to slide1_modified.png.
   - Run `ppt_visual_diff(slide1.png, slide1_modified.png)` -> assert similarity $< 100\%$ and changed bounding box coordinates tightly encapsulate the title's shift.
