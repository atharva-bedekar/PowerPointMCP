#!/usr/bin/env python3
"""Standalone CLI utility to inspect PowerPoint (.pptx) presentations, slides, and shapes."""

import argparse
import json
from pathlib import Path
import sys
from typing import Any, Dict

from powerpoint_mcp.pptx.inspector import (
    inspect_presentation,
    inspect_shape,
    inspect_slide,
)


def _format_ascii_tree(data: Dict[str, Any], verbose: bool = False) -> str:
    """Format presentation or slide model into a clean human-readable ASCII summary."""
    lines = []

    # Presentation level
    if "slide_count" in data and "layouts" in data:
        lines.append("=" * 70)
        lines.append(f"PRESENTATION: {data.get('path', '<in-memory>')}")
        lines.append(f"Dimensions: {data.get('width_inches', 0):.2f}\" x {data.get('height_inches', 0):.2f}\"")
        lines.append(f"Total Slides: {data.get('slide_count', 0)}")
        lines.append(f"Layouts: {', '.join(data.get('layouts', []))}")
        lines.append("-" * 70)
        lines.append("SLIDES SUMMARY:")
        for s in data.get("slides", []):
            if isinstance(s, dict):
                lines.append(f"  Slide {s.get('slide_number')}: \"{s.get('title') or '(Untitled)'}\" ({s.get('shape_count', 0)} shapes, layout: {s.get('layout_name')})")
        lines.append("=" * 70)

    # Slide level
    elif "slide_number" in data and "shapes" in data:
        lines.append("=" * 70)
        lines.append(f"SLIDE {data.get('slide_number')}: \"{data.get('title') or '(Untitled)'}\"")
        lines.append(f"Layout: {data.get('layout_name')} | Shapes: {data.get('shape_count', 0)} | Dimensions: {data.get('width_inches', 0):.2f}\" x {data.get('height_inches', 0):.2f}\"")
        lines.append("-" * 70)
        for s in data.get("shapes", []):
            bbox = s.get("bbox", {})
            lines.append(f"  [ID {s.get('shape_id'):3d}] {s.get('name')} ({s.get('semantic_role')}, {s.get('shape_type')})")
            lines.append(f"         Pos: (x={bbox.get('left_inches', 0):.2f}\", y={bbox.get('top_inches', 0):.2f}\") | Size: ({bbox.get('width_inches', 0):.2f}\" x {bbox.get('height_inches', 0):.2f}\")")
            tf = s.get("text_frame")
            if tf and tf.get("text"):
                preview = tf["text"].replace("\n", " ").strip()
                if len(preview) > 60:
                    preview = preview[:57] + "..."
                lines.append(f"         Text: \"{preview}\"")
            if verbose and s.get("properties"):
                lines.append(f"         Properties: {json.dumps(s.get('properties'))}")
        lines.append("=" * 70)

    # Shape level
    elif "shape_id" in data:
        lines.append("=" * 70)
        lines.append(f"SHAPE ID {data.get('shape_id')}: {data.get('name')}")
        lines.append(f"Type: {data.get('shape_type')} | Semantic Role: {data.get('semantic_role')}")
        bbox = data.get("bbox", {})
        lines.append(f"Bounding Box: x={bbox.get('left_inches', 0):.2f}\", y={bbox.get('top_inches', 0):.2f}\", w={bbox.get('width_inches', 0):.2f}\", h={bbox.get('height_inches', 0):.2f}\"")
        tf = data.get("text_frame")
        if tf:
            lines.append(f"Text Content:\n{tf.get('text', '')}")
            lines.append(f"Paragraphs: {len(tf.get('paragraphs', []))}")
        if data.get("fill"):
            lines.append(f"Fill: {data.get('fill')}")
        if data.get("line"):
            lines.append(f"Line: {data.get('line')}")
        if verbose and data.get("properties"):
            lines.append(f"Raw Properties:\n{json.dumps(data.get('properties'), indent=2)}")
        lines.append("=" * 70)

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspect PowerPoint (.pptx) presentation metadata, slides, and shape trees."
    )
    parser.add_argument("presentation", type=str, help="Path to .pptx presentation file")
    parser.add_argument("--slide", type=int, default=None, help="1-indexed slide number to inspect")
    parser.add_argument("--shape", type=int, default=None, help="Target shape ID on specified slide")
    parser.add_argument("--json", action="store_true", help="Output raw JSON format instead of ASCII tree")
    parser.add_argument("--verbose", action="store_true", help="Include extended properties and metadata")

    args = parser.parse_args()

    file_path = Path(args.presentation).resolve()
    if not file_path.exists():
        sys.stderr.write(f"Error: Presentation file not found: {file_path}\n")
        return 1

    try:
        if args.shape is not None:
            if args.slide is None:
                sys.stderr.write("Error: --slide <number> is required when inspecting a specific --shape <ID>\n")
                return 1
            shape_model = inspect_shape(str(file_path), args.slide, args.shape)
            out_data = shape_model.to_dict()
        elif args.slide is not None:
            slide_model = inspect_slide(str(file_path), args.slide)
            out_data = slide_model.to_dict()
        else:
            prs_model = inspect_presentation(str(file_path))
            out_data = prs_model.to_dict()

        if args.json:
            print(json.dumps(out_data, indent=2))
        else:
            print(_format_ascii_tree(out_data, verbose=args.verbose))

        return 0

    except Exception as exc:
        sys.stderr.write(f"Error parsing presentation: {exc}\n")
        return 2


if __name__ == "__main__":
    sys.exit(main())
