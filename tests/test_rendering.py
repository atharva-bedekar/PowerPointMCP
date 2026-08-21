"""Comprehensive tests for slide rendering, COM automation, LibreOffice fallback, visual diffing, and slide comparison."""

import os
from pathlib import Path
import sys
import tempfile

import numpy as np
from PIL import Image, ImageDraw
import pytest

from powerpoint_mcp.models.shape import (
    BoundingBox,
    ParagraphModel,
    SemanticRole,
    ShapeModel,
    ShapeType,
    TextFrameModel,
    TextRunModel,
    TextStyle,
)
from powerpoint_mcp.models.slide import SlideModel
from powerpoint_mcp.pptx.inspector import inspect_slide
from powerpoint_mcp.rendering import (
    BaseRenderer,
    LibreOfficeRenderer,
    NullRenderer,
    PowerPointRenderer,
    SlideComparisonResult,
    VisualDiffResult,
    compare_slides,
    get_available_renderer,
    visual_diff,
)


@pytest.fixture
def synthetic_sample_path():
    """Return path to synthetic test presentation."""
    path = Path(__file__).resolve().parent / "fixtures" / "synthetic_sample.pptx"
    assert path.exists(), f"Synthetic presentation fixture missing at {path}"
    return path


@pytest.fixture
def sample_images(tmp_path):
    """Generate sample test images for visual diff tests."""
    w, h = 800, 600

    # Image A: White canvas with blue box at (100, 100, 300, 200) and text
    img_a = Image.new("RGB", (w, h), color=(255, 255, 255))
    draw_a = ImageDraw.Draw(img_a)
    draw_a.rectangle([100, 100, 300, 200], fill=(41, 128, 185), outline=(0, 0, 0), width=2)
    path_a = tmp_path / "img_a.png"
    img_a.save(str(path_a), format="PNG")

    # Image B (Identical to A)
    path_b_identical = tmp_path / "img_b_identical.png"
    img_a.save(str(path_b_identical), format="PNG")

    # Image C (Modified: box moved to (200, 100, 400, 200) + green circle)
    img_c = Image.new("RGB", (w, h), color=(255, 255, 255))
    draw_c = ImageDraw.Draw(img_c)
    draw_c.rectangle([200, 100, 400, 200], fill=(41, 128, 185), outline=(0, 0, 0), width=2)
    draw_c.ellipse([500, 300, 600, 400], fill=(39, 174, 96))
    path_c = tmp_path / "img_c_modified.png"
    img_c.save(str(path_c), format="PNG")

    # Image D (Different dimensions: 400x300)
    img_d = Image.new("RGB", (400, 300), color=(240, 240, 240))
    path_d = tmp_path / "img_d_small.png"
    img_d.save(str(path_d), format="PNG")

    return {
        "path_a": path_a,
        "path_b_identical": path_b_identical,
        "path_c_modified": path_c,
        "path_d_small": path_d,
    }


# =============================================================================
# 1. Renderer Detection & Configuration Tests
# =============================================================================


class TestRendererDetection:
    """Tests for renderer selection and discovery mechanism."""

    def test_get_available_renderer_auto(self):
        renderer = get_available_renderer(preferred="auto")
        assert isinstance(renderer, BaseRenderer)
        assert isinstance(renderer.renderer_name, str)
        assert isinstance(renderer.is_available, bool)

    def test_get_available_renderer_powerpoint(self):
        renderer = get_available_renderer(preferred="powerpoint")
        assert isinstance(renderer, PowerPointRenderer)
        assert renderer.renderer_name == "powerpoint"

    def test_get_available_renderer_libreoffice(self):
        renderer = get_available_renderer(preferred="libreoffice")
        assert isinstance(renderer, LibreOfficeRenderer)
        assert renderer.renderer_name == "libreoffice"

    def test_get_available_renderer_none(self):
        renderer = get_available_renderer(preferred="none")
        assert isinstance(renderer, NullRenderer)
        assert renderer.renderer_name == "none"
        assert not renderer.is_available

    def test_null_renderer_raises_on_render(self, tmp_path):
        renderer = NullRenderer()
        with pytest.raises(RuntimeError, match="No presentation renderer is available"):
            renderer.render_slide("dummy.pptx", 1, tmp_path / "out.png")

        with pytest.raises(RuntimeError, match="No presentation renderer is available"):
            renderer.render_presentation("dummy.pptx", tmp_path / "out_dir")

    def test_env_var_override(self, monkeypatch):
        monkeypatch.setenv("PPT_RENDERER", "none")
        renderer = get_available_renderer("auto")
        assert isinstance(renderer, NullRenderer)

        monkeypatch.setenv("PPT_RENDERER", "libreoffice")
        renderer = get_available_renderer("auto")
        assert isinstance(renderer, LibreOfficeRenderer)

    def test_renderer_info_metadata(self):
        renderer = get_available_renderer("auto")
        info = renderer.get_renderer_info()
        assert "renderer_name" in info
        assert "is_available" in info
        assert "platform" in info


# =============================================================================
# 2. PowerPoint COM Renderer Tests (Windows)
# =============================================================================


class TestPowerPointCOM:
    """Tests for native Microsoft PowerPoint COM automation on Windows."""

    @pytest.mark.skipif(
        sys.platform != "win32" or not PowerPointRenderer().is_available,
        reason="PowerPoint COM is not available on this platform",
    )
    def test_render_single_slide_com(self, synthetic_sample_path, tmp_path):
        renderer = PowerPointRenderer()
        out_png = tmp_path / "slide_1.png"

        result_path = renderer.render_slide(
            presentation_path=synthetic_sample_path,
            slide_number=1,
            output_path=out_png,
            width=1920,
            height=1080,
        )

        assert Path(result_path).exists()
        assert Path(result_path).stat().st_size > 0

        # Verify valid PNG file header
        with open(result_path, "rb") as f:
            header = f.read(8)
            assert header == b"\x89PNG\r\n\x1a\n", "Generated file is not a valid PNG image"

        # Verify image dimensions via Pillow
        with Image.open(result_path) as img:
            assert img.size == (1920, 1080)
            assert img.format == "PNG"

    @pytest.mark.skipif(
        sys.platform != "win32" or not PowerPointRenderer().is_available,
        reason="PowerPoint COM is not available on this platform",
    )
    def test_render_presentation_all_slides_com(self, synthetic_sample_path, tmp_path):
        renderer = PowerPointRenderer()
        out_dir = tmp_path / "rendered_deck"

        rendered_files = renderer.render_presentation(
            presentation_path=synthetic_sample_path,
            output_dir=out_dir,
            width=1280,
            height=720,
        )

        assert len(rendered_files) == 3
        for idx, slide_path in enumerate(rendered_files, start=1):
            p = Path(slide_path)
            assert p.exists()
            assert p.name == f"slide_{idx}.png"
            with Image.open(p) as img:
                assert img.size == (1280, 720)

    @pytest.mark.skipif(
        sys.platform != "win32" or not PowerPointRenderer().is_available,
        reason="PowerPoint COM is not available on this platform",
    )
    def test_slide_number_out_of_bounds_com(self, synthetic_sample_path, tmp_path):
        renderer = PowerPointRenderer()

        with pytest.raises(IndexError):
            renderer.render_slide(synthetic_sample_path, 0, tmp_path / "err.png")

        with pytest.raises(IndexError, match="out of range"):
            renderer.render_slide(synthetic_sample_path, 999, tmp_path / "err.png")

    def test_render_nonexistent_presentation(self, tmp_path):
        renderer = PowerPointRenderer()
        if renderer.is_available:
            with pytest.raises(FileNotFoundError):
                renderer.render_slide("nonexistent_file.pptx", 1, tmp_path / "err.png")


# =============================================================================
# 3. LibreOffice Fallback Renderer Tests
# =============================================================================


class TestLibreOfficeRenderer:
    """Tests for LibreOffice headless renderer detection and error handling."""

    def test_libreoffice_properties(self):
        renderer = LibreOfficeRenderer()
        assert renderer.renderer_name == "libreoffice"
        info = renderer.get_renderer_info()
        assert "executable_path" in info

    def test_custom_executable_path(self, tmp_path):
        custom_bin = tmp_path / "fake_soffice.exe"
        custom_bin.touch()
        renderer = LibreOfficeRenderer(executable_path=str(custom_bin))
        assert renderer.is_available
        assert renderer._find_executable() == str(custom_bin.resolve())

    def test_render_unavailable_raises_runtime_error(self, tmp_path):
        renderer = LibreOfficeRenderer(executable_path=r"C:\fake\nonexistent\soffice.exe")
        assert not renderer.is_available
        with pytest.raises(RuntimeError, match="LibreOffice executable"):
            renderer.render_slide("dummy.pptx", 1, tmp_path / "out.png")


# =============================================================================
# 4. Visual Diffing & Image Comparison Tests
# =============================================================================


class TestVisualDiff:
    """Tests for image_diff module: pixel metrics, bounding box clustering, overlay."""

    def test_identical_images(self, sample_images):
        result = visual_diff(
            sample_images["path_a"],
            sample_images["path_b_identical"],
            threshold=25,
        )

        assert isinstance(result, VisualDiffResult)
        assert result.similarity_percentage == 100.0
        assert result.pixel_diff_count == 0
        assert result.mse == 0.0
        assert result.psnr == float("inf")
        assert len(result.changed_bounding_boxes) == 0
        assert result.is_identical is True

    def test_modified_images_metrics_and_boxes(self, sample_images, tmp_path):
        diff_png = tmp_path / "diff_output.png"

        result = visual_diff(
            sample_images["path_a"],
            sample_images["path_c_modified"],
            diff_output_path=diff_png,
            threshold=25,
            block_size=32,
        )

        assert result.similarity_percentage < 100.0
        assert result.similarity_percentage > 80.0
        assert result.pixel_diff_count > 0
        assert result.mse > 0.0
        assert 0.0 < result.psnr < 100.0
        assert result.is_identical is False
        assert len(result.changed_bounding_boxes) >= 1

        # Check bounding box keys
        for box in result.changed_bounding_boxes:
            assert "x" in box
            assert "y" in box
            assert "width" in box
            assert "height" in box
            assert "right" in box
            assert "bottom" in box
            assert box["width"] > 0
            assert box["height"] > 0

        # Verify diff artifact image
        assert diff_png.exists()
        assert diff_png.stat().st_size > 0
        with Image.open(diff_png) as dimg:
            assert dimg.size == (800, 600)
            # Verify magenta highlights exist in diff image array
            d_arr = np.array(dimg)
            magenta_mask = (
                (d_arr[:, :, 0] == 255) & (d_arr[:, :, 1] == 0) & (d_arr[:, :, 2] == 255)
            )
            assert np.any(magenta_mask), "Diff image should contain magenta (#FF00FF) highlights"

    def test_visual_diff_dimension_mismatch_resizes(self, sample_images):
        # image_a is 800x600, image_d is 400x300
        result = visual_diff(
            sample_images["path_a"],
            sample_images["path_d_small"],
            threshold=25,
        )
        assert result.total_pixels == 800 * 600
        assert result.similarity_percentage >= 0.0

    def test_visual_diff_missing_file_raises(self, sample_images):
        with pytest.raises(FileNotFoundError):
            visual_diff("missing_image_1.png", sample_images["path_a"])

        with pytest.raises(FileNotFoundError):
            visual_diff(sample_images["path_a"], "missing_image_2.png")

    def test_visual_diff_result_to_dict(self, sample_images):
        result = visual_diff(
            sample_images["path_a"],
            sample_images["path_b_identical"],
        )
        d = result.to_dict()
        assert d["similarity_percentage"] == 100.0
        assert d["pixel_diff_count"] == 0
        assert d["is_identical"] is True
        assert d["psnr"] == "Infinity"


# =============================================================================
# 5. Slide Model & Visual Comparison Tests
# =============================================================================


class TestCompareSlides:
    """Tests for compare_slides combining geometric inspection, typography, and visual diffs."""

    def test_compare_identical_slides(self, synthetic_sample_path):
        slide1 = inspect_slide(synthetic_sample_path, 1)
        res = compare_slides(slide1, slide1)

        assert isinstance(res, SlideComparisonResult)
        assert res.slide_a_number == 1
        assert res.slide_b_number == 1
        assert res.geometric_match_score == 100.0
        assert res.overall_similarity_score == 100.0
        assert res.dimension_match is True
        assert res.layout_match is True
        assert len(res.unmatched_shapes_a) == 0
        assert len(res.unmatched_shapes_b) == 0
        assert len(res.layout_differences) == 0
        assert len(res.typography_differences) == 0
        assert "identical" in res.summary.lower()

    def test_compare_different_slides(self, synthetic_sample_path):
        slide1 = inspect_slide(synthetic_sample_path, 1)
        slide2 = inspect_slide(synthetic_sample_path, 2)
        res = compare_slides(slide1, slide2)

        assert res.slide_a_number == 1
        assert res.slide_b_number == 2
        assert res.shape_count_a == len(slide1.shapes)
        assert res.shape_count_b == len(slide2.shapes)
        assert res.geometric_match_score < 100.0
        assert len(res.summary) > 0

    def test_compare_slides_with_rendered_images(self, synthetic_sample_path, sample_images, tmp_path):
        slide1 = inspect_slide(synthetic_sample_path, 1)
        diff_out = tmp_path / "slide_diff.png"

        res = compare_slides(
            slide1,
            slide1,
            slide_a_img_path=sample_images["path_a"],
            slide_b_img_path=sample_images["path_c_modified"],
            diff_output_path=diff_out,
        )

        assert res.visual_diff is not None
        assert res.visual_diff.similarity_percentage < 100.0
        assert res.overall_similarity_score < 100.0
        assert diff_out.exists()

    def test_compare_slides_detects_layout_and_typography_shift(self):
        # Construct synthetic slide A
        box_a = BoundingBox(left_emu=914400, top_emu=914400, width_emu=1828800, height_emu=914400)
        style_a = TextStyle(font_name="Calibri", font_size_pt=18.0, bold=False, color_rgb="1F497D")
        run_a = TextRunModel(text="Original Header", style=style_a)
        tf_a = TextFrameModel(text="Original Header", paragraphs=[ParagraphModel(text="Original Header", runs=[run_a])])
        shape_a = ShapeModel(
            shape_id=1,
            name="Title 1",
            shape_type=ShapeType.TEXT_BOX,
            semantic_role=SemanticRole.TITLE,
            bbox=box_a,
            text_frame=tf_a,
        )
        slide_a = SlideModel(slide_number=1, slide_id=256, layout_name="Title", shapes=[shape_a])

        # Construct synthetic slide B with moved position and modified typography
        box_b = BoundingBox(left_emu=1828800, top_emu=914400, width_emu=1828800, height_emu=914400)  # moved right 1 inch
        style_b = TextStyle(font_name="Arial", font_size_pt=24.0, bold=True, color_rgb="FF0000")
        run_b = TextRunModel(text="Updated Header", style=style_b)
        tf_b = TextFrameModel(text="Updated Header", paragraphs=[ParagraphModel(text="Updated Header", runs=[run_b])])
        shape_b = ShapeModel(
            shape_id=1,
            name="Title 1",
            shape_type=ShapeType.TEXT_BOX,
            semantic_role=SemanticRole.TITLE,
            bbox=box_b,
            text_frame=tf_b,
        )
        slide_b = SlideModel(slide_number=1, slide_id=256, layout_name="Title", shapes=[shape_b])

        res = compare_slides(slide_a, slide_b)

        assert len(res.layout_differences) >= 1
        assert len(res.typography_differences) >= 1

        # Check geometry shift detected
        geo_diff = [d for d in res.layout_differences if d["type"] == "shape_geometry_shift"]
        assert len(geo_diff) == 1
        assert geo_diff[0]["delta_x_inches"] == 1.0

        # Check typography changes detected
        typo_change = [d for d in res.typography_differences if d["type"] == "typography_style_change"]
        assert len(typo_change) == 1
        diffs = typo_change[0]["differences"]
        assert diffs["font_name"]["a"] == "Calibri"
        assert diffs["font_name"]["b"] == "Arial"
        assert diffs["bold"]["a"] is False
        assert diffs["bold"]["b"] is True

    def test_slide_comparison_result_to_dict(self, synthetic_sample_path):
        slide1 = inspect_slide(synthetic_sample_path, 1)
        res = compare_slides(slide1, slide1)
        d = res.to_dict()

        assert d["slide_a_number"] == 1
        assert d["slide_b_number"] == 1
        assert d["geometric_match_score"] == 100.0
        assert d["dimension_match"] is True
        assert d["matched_shape_count"] == len(slide1.shapes)
