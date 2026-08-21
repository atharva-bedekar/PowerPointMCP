"""Comprehensive test suite for PPTX geometry engine (alignment, distribution, equalization, overlap detection, boundary validation)."""

import pytest
from pptx import Presentation
from pptx.util import Inches

from powerpoint_mcp.models.shape import (
    AlignmentType,
    BoundingBox,
    DistributionMode,
    EMU_PER_INCH,
    ShapeModel,
    ShapeType,
    SemanticRole,
    SpacingMode,
    inches_to_emu,
    emu_to_inches,
)
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.geometry import (
    align_shapes,
    distribute_shapes,
    equalize_dimensions,
    check_bounding_box_collision,
    calculate_overlap_box,
    calculate_overlap_area,
    calculate_overlap_area_sq_inches,
    detect_slide_overlaps,
    detect_off_slide_shapes,
)
from powerpoint_mcp.pptx.inspector import inspect_slide


class TestShapeAlignment:
    """Test 6-axis shape alignment algorithm across models, bounding boxes, and python-pptx shapes."""

    def test_align_left_models(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 1.0)
        b2 = BoundingBox.from_inches(3.0, 2.0, 1.5, 1.0)
        b3 = BoundingBox.from_inches(0.5, 4.0, 2.5, 1.0)

        boxes = [b1, b2, b3]
        aligned = align_shapes(boxes, AlignmentType.LEFT)

        assert len(aligned) == 3
        # Leftmost is b3 (0.5 inches = 457200 EMU)
        min_left = inches_to_emu(0.5)
        for b in aligned:
            assert b.left_emu == min_left
            assert b.left_inches == 0.5

    def test_align_left_with_reference(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 1.0)
        b2 = BoundingBox.from_inches(3.0, 2.0, 1.5, 1.0)
        ref = BoundingBox.from_inches(4.0, 0.0, 1.0, 1.0)

        align_shapes([b1, b2], AlignmentType.LEFT, reference_shape=ref)
        assert b1.left_inches == 4.0
        assert b2.left_inches == 4.0

    def test_align_center_models(self):
        # b1 center_x = 1.0 + 1.0 = 2.0
        # b2 center_x = 4.0 + 1.0 = 5.0
        # Average center_x = (2.0 + 5.0) / 2 = 3.5 inches
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 1.0)
        b2 = BoundingBox.from_inches(4.0, 2.0, 2.0, 1.0)

        align_shapes([b1, b2], "center")
        assert b1.center_x_inches == pytest.approx(3.5, abs=1e-4)
        assert b2.center_x_inches == pytest.approx(3.5, abs=1e-4)
        assert b1.left_inches == pytest.approx(2.5, abs=1e-4)
        assert b2.left_inches == pytest.approx(2.5, abs=1e-4)

    def test_align_right_models(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 1.0) # right = 3.0
        b2 = BoundingBox.from_inches(2.0, 2.0, 3.0, 1.0) # right = 5.0
        # Max right = 5.0

        align_shapes([b1, b2], AlignmentType.RIGHT)
        assert b1.right_inches == 5.0
        assert b2.right_inches == 5.0
        assert b1.left_inches == 3.0 # 5.0 - 2.0
        assert b2.left_inches == 2.0 # 5.0 - 3.0

    def test_align_top_middle_bottom(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 1.0, 2.0) # top=1.0, bot=3.0, cy=2.0
        b2 = BoundingBox.from_inches(3.0, 3.0, 1.0, 4.0) # top=3.0, bot=7.0, cy=5.0

        # Top alignment: min top is 1.0
        align_shapes([b1, b2], AlignmentType.TOP)
        assert b1.top_inches == 1.0
        assert b2.top_inches == 1.0

        # Bottom alignment: max bottom is 1.0+2.0=3.0 and 1.0+4.0=5.0 -> max is 5.0
        align_shapes([b1, b2], AlignmentType.BOTTOM)
        assert b1.bottom_inches == 5.0
        assert b1.top_inches == 3.0 # 5.0 - 2.0
        assert b2.bottom_inches == 5.0
        assert b2.top_inches == 1.0 # 5.0 - 4.0

        # Middle alignment: with reference
        ref = BoundingBox.from_inches(0.0, 2.0, 1.0, 2.0) # cy = 3.0
        align_shapes([b1, b2], AlignmentType.MIDDLE, reference_shape=ref)
        assert b1.center_y_inches == 3.0
        assert b2.center_y_inches == 3.0
        assert b1.top_inches == 2.0 # 3.0 - 1.0
        assert b2.top_inches == 1.0 # 3.0 - 2.0

    def test_align_python_pptx_shapes_in_place(self, temp_presentation):
        slide = temp_presentation.slides[0]
        shapes = list(slide.shapes)[:3]
        if len(shapes) >= 2:
            align_shapes(shapes, AlignmentType.LEFT)
            left_0 = shapes[0].left
            for s in shapes:
                assert s.left == left_0


class TestShapeDistribution:
    """Test horizontal and vertical distribution with equal gaps and equal centers."""

    def test_distribute_horizontal_equal_gaps(self):
        # 3 boxes of width 1.0 at x=0.0, x=3.0, x=8.0
        # Total span = (8.0 + 1.0) - 0.0 = 9.0 inches
        # Total widths = 3.0 inches
        # Total gap = 6.0 inches -> gap per space = 6.0 / 2 = 3.0 inches
        # Positions should be:
        # b0: left=0.0, right=1.0
        # b1: left=1.0 + 3.0 = 4.0, right=5.0
        # b2: left=5.0 + 3.0 = 8.0, right=9.0
        b0 = BoundingBox.from_inches(0.0, 1.0, 1.0, 1.0)
        b1 = BoundingBox.from_inches(3.0, 1.0, 1.0, 1.0)
        b2 = BoundingBox.from_inches(8.0, 1.0, 1.0, 1.0)

        distributed = distribute_shapes([b1, b0, b2], DistributionMode.HORIZONTAL, SpacingMode.EQUAL_GAPS)

        assert distributed[0].left_inches == 0.0
        assert distributed[1].left_inches == 4.0
        assert distributed[2].left_inches == 8.0

    def test_distribute_horizontal_equal_centers(self):
        # 3 boxes of varying widths
        # b0: w=1.0 at 0.0 -> center = 0.5
        # b1: w=2.0 at 4.0
        # b2: w=1.0 at 6.0 -> center = 6.5
        # Total center span = 6.5 - 0.5 = 6.0 -> step = 3.0
        # b1 target center = 3.5 -> left = 3.5 - 2.0 / 2 = 2.5
        b0 = BoundingBox.from_inches(0.0, 1.0, 1.0, 1.0)
        b1 = BoundingBox.from_inches(4.0, 1.0, 2.0, 1.0)
        b2 = BoundingBox.from_inches(6.0, 1.0, 1.0, 1.0)

        distributed = distribute_shapes([b0, b1, b2], "horizontal", "equal_centers")

        assert distributed[0].center_x_inches == 0.5
        assert distributed[1].center_x_inches == 3.5
        assert distributed[1].left_inches == 2.5
        assert distributed[2].center_x_inches == 6.5

    def test_distribute_vertical_equal_gaps(self):
        # 3 boxes of height 1.0 at y=1.0, y=5.0, y=7.0
        # Total span = (7.0 + 1.0) - 1.0 = 7.0
        # Total heights = 3.0 -> Total gap = 4.0 -> gap = 2.0
        # b0: y=1.0
        # b1: y=1.0 + 1.0 + 2.0 = 4.0
        # b2: y=4.0 + 1.0 + 2.0 = 7.0
        b0 = BoundingBox.from_inches(1.0, 1.0, 1.0, 1.0)
        b1 = BoundingBox.from_inches(1.0, 5.0, 1.0, 1.0)
        b2 = BoundingBox.from_inches(1.0, 7.0, 1.0, 1.0)

        distribute_shapes([b0, b1, b2], DistributionMode.VERTICAL, SpacingMode.EQUAL_GAPS)

        assert b0.top_inches == 1.0
        assert b1.top_inches == 4.0
        assert b2.top_inches == 7.0

    def test_distribute_vertical_equal_centers(self):
        b0 = BoundingBox.from_inches(1.0, 0.0, 1.0, 2.0) # cy = 1.0
        b1 = BoundingBox.from_inches(1.0, 4.0, 1.0, 2.0)
        b2 = BoundingBox.from_inches(1.0, 8.0, 1.0, 2.0) # cy = 9.0
        # step = (9.0 - 1.0)/2 = 4.0 -> target cy = 5.0 -> top = 5.0 - 1.0 = 4.0

        distribute_shapes([b0, b1, b2], DistributionMode.VERTICAL, SpacingMode.EQUAL_CENTERS)
        assert b1.center_y_inches == 5.0
        assert b1.top_inches == 4.0

    def test_distribute_fewer_than_three_shapes(self):
        b0 = BoundingBox.from_inches(1.0, 1.0, 1.0, 1.0)
        b1 = BoundingBox.from_inches(5.0, 1.0, 1.0, 1.0)

        res = distribute_shapes([b0, b1], DistributionMode.HORIZONTAL)
        assert len(res) == 2
        assert res[0].left_inches == 1.0
        assert res[1].left_inches == 5.0


class TestDimensionEqualization:
    """Test width and height equalization strategies."""

    def test_equalize_dimensions_first_mode(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 3.0)
        b2 = BoundingBox.from_inches(2.0, 2.0, 4.0, 1.0)
        b3 = BoundingBox.from_inches(3.0, 3.0, 1.0, 5.0)

        equalize_dimensions([b1, b2, b3], equalize_width=True, equalize_height=True, mode="first")

        for b in [b1, b2, b3]:
            assert b.width_inches == 2.0
            assert b.height_inches == 3.0

    def test_equalize_dimensions_max_and_min(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 3.0)
        b2 = BoundingBox.from_inches(2.0, 2.0, 5.0, 1.0)

        equalize_dimensions([b1, b2], equalize_width=True, equalize_height=False, mode="max")
        assert b1.width_inches == 5.0
        assert b2.width_inches == 5.0
        assert b1.height_inches == 3.0
        assert b2.height_inches == 1.0

        equalize_dimensions([b1, b2], equalize_width=False, equalize_height=True, mode="min")
        assert b1.height_inches == 1.0
        assert b2.height_inches == 1.0

    def test_equalize_explicit_target_inches(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 2.0)
        b2 = BoundingBox.from_inches(2.0, 2.0, 3.0, 3.0)

        equalize_dimensions([b1, b2], target_width_inches=4.5, target_height_inches=2.25)
        assert b1.width_inches == 4.5
        assert b2.width_inches == 4.5
        assert b1.height_inches == 2.25
        assert b2.height_inches == 2.25


class TestCollisionAndOverlapMath:
    """Test geometric collision detection, intersection box calculation, and slide overlap analysis."""

    def test_check_bounding_box_collision(self):
        # Overlapping boxes
        b1 = BoundingBox.from_inches(1.0, 1.0, 3.0, 3.0) # (1, 1) to (4, 4)
        b2 = BoundingBox.from_inches(2.0, 2.0, 3.0, 3.0) # (2, 2) to (5, 5)

        assert check_bounding_box_collision(b1, b2) is True

        # Non-overlapping (separated)
        b3 = BoundingBox.from_inches(5.0, 5.0, 2.0, 2.0)
        assert check_bounding_box_collision(b1, b3) is False

        # Edge touching (zero area overlap)
        b4 = BoundingBox.from_inches(4.0, 1.0, 2.0, 3.0) # left=4.0 equals b1 right=4.0
        assert check_bounding_box_collision(b1, b4, tolerance_emu=0) is False

    def test_calculate_overlap_box_and_area(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 4.0, 3.0) # x: 1..5, y: 1..4
        b2 = BoundingBox.from_inches(3.0, 2.0, 4.0, 4.0) # x: 3..7, y: 2..6

        overlap = calculate_overlap_box(b1, b2)
        assert overlap is not None
        assert overlap.left_inches == 3.0
        assert overlap.top_inches == 2.0
        assert overlap.width_inches == 2.0 # 5 - 3
        assert overlap.height_inches == 2.0 # 4 - 2

        area_sq_in = calculate_overlap_area_sq_inches(b1, b2)
        assert area_sq_in == pytest.approx(4.0, abs=1e-4)

    def test_detect_slide_overlaps_synthetic_deck(self, synthetic_deck_path):
        """Slide 3 of the synthetic deck contains deliberate overlapping shapes."""
        slide_3 = inspect_slide(synthetic_deck_path, 3)
        overlaps = detect_slide_overlaps(slide_3, min_overlap_area_sq_in=0.01)

        assert len(overlaps) >= 1
        overlap = overlaps[0]
        assert "shape_1_id" in overlap
        assert "shape_2_id" in overlap
        assert overlap["overlap_area_sq_in"] > 0.01
        assert overlap["overlap_bbox"] is not None


class TestBoundaryValidation:
    """Test detection of off-canvas and clipped shapes."""

    def test_detect_off_slide_shapes(self):
        slide_width = 10.0
        slide_height = 5.625

        # Normal contained shape
        normal_box = BoundingBox.from_inches(1.0, 1.0, 2.0, 1.0)

        # Off-left shape
        off_left = BoundingBox.from_inches(-0.5, 1.0, 2.0, 1.0)

        # Off-right shape (starts at 9.0, width 2.0 -> right = 11.0, breach = 1.0)
        off_right = BoundingBox.from_inches(9.0, 1.0, 2.0, 1.0)

        # Off-bottom shape (top = 5.0, height = 1.5 -> bottom = 6.5, breach = 0.875)
        off_bottom = BoundingBox.from_inches(1.0, 5.0, 2.0, 1.5)

        shapes = [
            ShapeModel(
                shape_id=1,
                name="Normal",
                shape_type=ShapeType.AUTO_SHAPE,
                semantic_role=SemanticRole.BODY,
                bbox=normal_box,
            ),
            ShapeModel(
                shape_id=2,
                name="OffLeft",
                shape_type=ShapeType.AUTO_SHAPE,
                semantic_role=SemanticRole.BODY,
                bbox=off_left,
            ),
            ShapeModel(
                shape_id=3,
                name="OffRight",
                shape_type=ShapeType.AUTO_SHAPE,
                semantic_role=SemanticRole.BODY,
                bbox=off_right,
            ),
            ShapeModel(
                shape_id=4,
                name="OffBottom",
                shape_type=ShapeType.AUTO_SHAPE,
                semantic_role=SemanticRole.BODY,
                bbox=off_bottom,
            ),
        ]

        breaches = detect_off_slide_shapes(shapes, slide_width, slide_height, tolerance_inches=0.01)

        assert len(breaches) == 3
        shape_ids = [b["shape_id"] for b in breaches]
        assert 2 in shape_ids
        assert 3 in shape_ids
        assert 4 in shape_ids

        left_breach = next(b for b in breaches if b["shape_id"] == 2)
        assert left_breach["breaches"]["left"] == pytest.approx(0.5, abs=1e-3)

        right_breach = next(b for b in breaches if b["shape_id"] == 3)
        assert right_breach["breaches"]["right"] == pytest.approx(1.0, abs=1e-3)

        bottom_breach = next(b for b in breaches if b["shape_id"] == 4)
        assert bottom_breach["breaches"]["bottom"] == pytest.approx(0.875, abs=1e-3)

    def test_detect_off_slide_shapes_synthetic_deck(self, synthetic_deck_path):
        """Slide 3 of the synthetic deck contains deliberate boundary breaches."""
        slide_3 = inspect_slide(synthetic_deck_path, 3)
        breaches = detect_off_slide_shapes(slide_3)

        assert len(breaches) >= 1
        for b in breaches:
            assert "shape_id" in b
            assert "breaches" in b
            assert len(b["breaches"]) > 0


class TestGeometryEdgeCases:
    """Test boundary and unexpected input scenarios for geometry algorithms."""

    def test_align_empty_and_single_shape(self):
        assert align_shapes([], AlignmentType.LEFT) == []
        b = BoundingBox.from_inches(1.0, 1.0, 2.0, 2.0)
        res = align_shapes([b], AlignmentType.CENTER)
        assert len(res) == 1
        assert res[0].left_inches == 1.0

    def test_align_invalid_mode_raises(self):
        b = BoundingBox.from_inches(1.0, 1.0, 2.0, 2.0)
        with pytest.raises(ValueError, match="Unknown alignment mode"):
            align_shapes([b, b], "diagonal")

    def test_distribute_invalid_modes_raise(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 1.0, 1.0)
        b2 = BoundingBox.from_inches(2.0, 1.0, 1.0, 1.0)
        b3 = BoundingBox.from_inches(3.0, 1.0, 1.0, 1.0)
        with pytest.raises(ValueError, match="Unknown distribution mode"):
            distribute_shapes([b1, b2, b3], "diagonal")

        with pytest.raises(ValueError, match="Unknown spacing mode"):
            distribute_shapes([b1, b2, b3], DistributionMode.HORIZONTAL, "random")

    def test_equalize_dimensions_avg_mode(self):
        b1 = BoundingBox.from_inches(1.0, 1.0, 2.0, 2.0)
        b2 = BoundingBox.from_inches(1.0, 1.0, 4.0, 4.0)
        equalize_dimensions([b1, b2], equalize_width=True, equalize_height=True, mode="avg")
        assert b1.width_inches == 3.0
        assert b2.width_inches == 3.0
        assert b1.height_inches == 3.0
        assert b2.height_inches == 3.0

    def test_collision_with_negative_coords(self):
        b1 = BoundingBox.from_inches(-2.0, -2.0, 3.0, 3.0) # (-2, -2) to (1, 1)
        b2 = BoundingBox.from_inches(0.0, 0.0, 2.0, 2.0) # (0, 0) to (2, 2)
        assert check_bounding_box_collision(b1, b2) is True
        overlap = calculate_overlap_box(b1, b2)
        assert overlap is not None
        assert overlap.left_inches == 0.0
        assert overlap.top_inches == 0.0
        assert overlap.width_inches == 1.0
        assert overlap.height_inches == 1.0

    def test_calculate_overlap_area_zero_when_disjoint(self):
        b1 = BoundingBox.from_inches(0.0, 0.0, 1.0, 1.0)
        b2 = BoundingBox.from_inches(2.0, 2.0, 1.0, 1.0)
        assert calculate_overlap_area(b1, b2) == 0
        assert calculate_overlap_area_sq_inches(b1, b2) == 0.0
        assert calculate_overlap_box(b1, b2) is None

