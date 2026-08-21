"""Deterministic visual image comparison, pixel diffing, bounding region clustering, and overlay generation."""

from collections import deque
from dataclasses import dataclass, field
import math
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
from PIL import Image, ImageDraw


@dataclass
class VisualDiffResult:
    """Result of visual comparison between two rendered images."""

    similarity_percentage: float  # 0.0 to 100.0%
    pixel_diff_count: int  # Number of changed pixels exceeding threshold
    total_pixels: int  # Total pixels in compared image (width * height)
    mse: float  # Mean Squared Error across RGB channels
    psnr: float  # Peak Signal-to-Noise Ratio in dB (inf if identical)
    changed_bounding_boxes: List[Dict[str, int]] = field(default_factory=list)
    diff_image_path: Optional[str] = None
    threshold: int = 25
    is_identical: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Serialize result to a dictionary."""
        return {
            "similarity_percentage": round(self.similarity_percentage, 4),
            "pixel_diff_count": int(self.pixel_diff_count),
            "total_pixels": int(self.total_pixels),
            "mse": round(self.mse, 6) if not math.isinf(self.mse) else 0.0,
            "psnr": round(self.psnr, 2) if not math.isinf(self.psnr) else "Infinity",
            "changed_bounding_boxes": self.changed_bounding_boxes,
            "diff_image_path": self.diff_image_path,
            "threshold": self.threshold,
            "is_identical": self.is_identical,
        }


def _cluster_changed_regions(
    mask: np.ndarray, block_size: int = 32
) -> List[Dict[str, int]]:
    """Group changed mask pixels into connected bounding box regions using grid clustering.

    Args:
        mask: 2D boolean numpy array (H, W) where True represents a changed pixel.
        block_size: Grid cell size for grouping neighboring changes.

    Returns:
        List of bounding box dictionaries with keys: 'x', 'y', 'width', 'height', 'right', 'bottom'.
    """
    height, width = mask.shape
    if not np.any(mask):
        return []

    grid_h = (height + block_size - 1) // block_size
    grid_w = (width + block_size - 1) // block_size

    # 1. Build active block grid
    active_grid = np.zeros((grid_h, grid_w), dtype=bool)
    for gy in range(grid_h):
        y0 = gy * block_size
        y1 = min(y0 + block_size, height)
        for gx in range(grid_w):
            x0 = gx * block_size
            x1 = min(x0 + block_size, width)
            if np.any(mask[y0:y1, x0:x1]):
                active_grid[gy, gx] = True

    # 2. Find 8-connected components in active block grid
    visited = np.zeros((grid_h, grid_w), dtype=bool)
    components: List[List[Tuple[int, int]]] = []

    for gy in range(grid_h):
        for gx in range(grid_w):
            if active_grid[gy, gx] and not visited[gy, gx]:
                # BFS to find connected component
                component: List[Tuple[int, int]] = []
                queue = deque([(gy, gx)])
                visited[gy, gx] = True

                while queue:
                    curr_gy, curr_gx = queue.popleft()
                    component.append((curr_gy, curr_gx))

                    # 8-connectivity neighbors
                    for dgy in (-1, 0, 1):
                        for dgx in (-1, 0, 1):
                            if dgy == 0 and dgx == 0:
                                continue
                            ny, nx = curr_gy + dgy, curr_gx + dgx
                            if 0 <= ny < grid_h and 0 <= nx < grid_w:
                                if active_grid[ny, nx] and not visited[ny, nx]:
                                    visited[ny, nx] = True
                                    queue.append((ny, nx))

                components.append(component)

    # 3. Compute exact pixel bounding boxes for each component
    bounding_boxes: List[Dict[str, int]] = []
    for component in components:
        min_x = width
        max_x = 0
        min_y = height
        max_y = 0

        for b_gy, b_gx in component:
            y0 = b_gy * block_size
            y1 = min(y0 + block_size, height)
            x0 = b_gx * block_size
            x1 = min(x0 + block_size, width)

            sub_mask = mask[y0:y1, x0:x1]
            if np.any(sub_mask):
                y_idx, x_idx = np.where(sub_mask)
                min_x = min(min_x, x0 + int(np.min(x_idx)))
                max_x = max(max_x, x0 + int(np.max(x_idx)))
                min_y = min(min_y, y0 + int(np.min(y_idx)))
                max_y = max(max_y, y0 + int(np.max(y_idx)))

        if min_x <= max_x and min_y <= max_y:
            bounding_boxes.append({
                "x": int(min_x),
                "y": int(min_y),
                "width": int(max_x - min_x + 1),
                "height": int(max_y - min_y + 1),
                "right": int(max_x + 1),
                "bottom": int(max_y + 1),
            })

    # Sort top-to-bottom, left-to-right
    bounding_boxes.sort(key=lambda b: (b["y"], b["x"]))
    return bounding_boxes


def visual_diff(
    image_a_path: Union[str, Path],
    image_b_path: Union[str, Path],
    diff_output_path: Optional[Union[str, Path]] = None,
    threshold: int = 25,
    block_size: int = 32,
) -> VisualDiffResult:
    """Compare two slide images, compute pixel similarity metrics, detect changed regions, and generate diff overlay.

    Args:
        image_a_path: Path to baseline image (e.g. before edit).
        image_b_path: Path to comparison image (e.g. after edit).
        diff_output_path: Optional path to write the visual diff image (PNG).
        threshold: Per-channel pixel difference threshold (0-255) to qualify as a change (default: 25).
        block_size: Grid block size in pixels for clustering changed regions (default: 32).

    Returns:
        VisualDiffResult containing similarity percentage, changed pixel count, MSE, PSNR,
        list of changed bounding boxes, and optional diff image path.
    """
    path_a = Path(image_a_path).resolve()
    path_b = Path(image_b_path).resolve()

    if not path_a.exists():
        raise FileNotFoundError(f"Baseline image not found: {path_a}")
    if not path_b.exists():
        raise FileNotFoundError(f"Comparison image not found: {path_b}")

    img_a = Image.open(str(path_a)).convert("RGB")
    img_b = Image.open(str(path_b)).convert("RGB")

    # If dimensions differ, resize image B to match image A
    if img_a.size != img_b.size:
        resample_filter = getattr(Image, "Resampling", Image).LANCZOS
        img_b = img_b.resize(img_a.size, resample=resample_filter)

    width, height = img_a.size
    total_pixels = width * height

    arr_a = np.asarray(img_a, dtype=np.float32)
    arr_b = np.asarray(img_b, dtype=np.float32)

    # 1. Pixel difference matrix
    diff_matrix = np.abs(arr_a - arr_b)  # Shape: (H, W, 3)
    max_channel_diff = np.max(diff_matrix, axis=2)  # Shape: (H, W)
    mask = max_channel_diff > float(threshold)

    pixel_diff_count = int(np.sum(mask))
    similarity_percentage = 100.0 * (1.0 - (float(pixel_diff_count) / float(total_pixels)))
    similarity_percentage = max(0.0, min(100.0, similarity_percentage))

    # 2. MSE & PSNR
    mse = float(np.mean((arr_a - arr_b) ** 2))
    if mse <= 1e-10:
        psnr = float("inf")
    else:
        psnr = float(20.0 * math.log10(255.0 / math.sqrt(mse)))

    is_identical = pixel_diff_count == 0

    # 3. Detect changed bounding regions
    changed_bounding_boxes = _cluster_changed_regions(mask, block_size=block_size)

    # 4. Generate Visual Diff Image if requested
    saved_diff_path: Optional[str] = None
    if diff_output_path is not None:
        out_p = Path(diff_output_path).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)

        # Create muted/lightened grayscale background from img_a
        gray_arr = np.asarray(img_a.convert("L"), dtype=np.float32)
        base_bg = (gray_arr * 0.65 + 255.0 * 0.35).astype(np.uint8)
        diff_arr = np.stack([base_bg, base_bg, base_bg], axis=2)

        # Highlight changed pixels with vivid magenta (#FF00FF)
        diff_arr[mask] = [255, 0, 255]

        diff_img = Image.fromarray(diff_arr, mode="RGB")
        draw = ImageDraw.Draw(diff_img)

        # Draw red bounding box rectangles around detected regions
        for box in changed_bounding_boxes:
            draw.rectangle(
                [box["x"], box["y"], box["right"] - 1, box["bottom"] - 1],
                outline=(255, 0, 0),
                width=2,
            )

        diff_img.save(str(out_p), format="PNG")
        saved_diff_path = str(out_p)

    return VisualDiffResult(
        similarity_percentage=similarity_percentage,
        pixel_diff_count=pixel_diff_count,
        total_pixels=total_pixels,
        mse=mse,
        psnr=psnr,
        changed_bounding_boxes=changed_bounding_boxes,
        diff_image_path=saved_diff_path,
        threshold=threshold,
        is_identical=is_identical,
    )
