#!/usr/bin/env python3
"""Standalone CLI utility to render PowerPoint (.pptx) slides to PNG images."""

import argparse
from pathlib import Path
import sys

from powerpoint_mcp.tools.rendering import (
    ppt_render_presentation,
    ppt_render_slide,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Render PowerPoint (.pptx) slides to high-resolution PNG images."
    )
    parser.add_argument("presentation", type=str, help="Path to .pptx presentation file")
    parser.add_argument("--slide", type=int, default=None, help="1-indexed slide number to render (default: all)")
    parser.add_argument("--output", type=str, default="./renders", help="Output directory path (default: ./renders)")
    parser.add_argument("--renderer", type=str, default="auto", choices=["auto", "powerpoint", "libreoffice", "mock", "pillow"], help="Renderer choice")
    parser.add_argument("--dpi", type=int, default=150, help="Render resolution DPI (default: 150)")

    args = parser.parse_args()

    file_path = Path(args.presentation).resolve()
    if not file_path.exists():
        sys.stderr.write(f"Error: Presentation file not found: {file_path}\n")
        return 1

    try:
        out_dir = Path(args.output).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        if args.slide is not None:
            res = ppt_render_slide(
                slide_number=args.slide,
                output_dir=str(out_dir),
                renderer=args.renderer,
                dpi=args.dpi,
                presentation_path=str(file_path),
            )
            if not res.get("success"):
                sys.stderr.write(f"Render Error: {res.get('message')}\n")
                return 1

            print(f"Rendered Slide {args.slide} -> {res.get('image_path')} (Renderer: {res.get('renderer')})")
            return 0

        else:
            res = ppt_render_presentation(
                output_dir=str(out_dir),
                renderer=args.renderer,
                dpi=args.dpi,
                presentation_path=str(file_path),
            )
            if not res.get("success"):
                sys.stderr.write(f"Render Error: {res.get('message')}\n")
                return 1

            print(f"Rendered {res.get('slide_count')} slide(s) using {res.get('renderer')}:")
            for item in res.get("rendered_slides", []):
                print(f"  Slide {item.get('slide_number')}: {item.get('image_path')}")
            return 0

    except Exception as exc:
        sys.stderr.write(f"Render Error: {exc}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
