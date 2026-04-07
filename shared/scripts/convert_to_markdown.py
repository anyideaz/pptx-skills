#!/usr/bin/env python3
"""
convert_to_markdown.py — Convert documents to Markdown using markitdown.

Usage:
    python convert_to_markdown.py <input_file> <output_file>

Exit codes:
    0  — success
    2  — file not found
    3  — conversion error
"""

import sys
import os


def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input_file> <output_file>", file=sys.stderr)
        sys.exit(3)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}", file=sys.stderr)
        sys.exit(2)

    try:
        from markitdown import MarkItDown
    except ImportError:
        print("ERROR: markitdown is not installed. Run setup_deps.sh or setup_deps.ps1.", file=sys.stderr)
        sys.exit(3)

    try:
        md = MarkItDown()
        result = md.convert(input_path)
        markdown_text = result.text_content
    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", file=sys.stderr)
        sys.exit(3)

    try:
        os.makedirs(os.path.dirname(output_path), exist_ok=True) if os.path.dirname(output_path) else None
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(markdown_text)
    except Exception as e:
        print(f"ERROR: Could not write output file: {e}", file=sys.stderr)
        sys.exit(3)

    ext = os.path.splitext(input_path)[1].lower()
    size = len(markdown_text)
    lines = markdown_text.count("\n") + 1
    print(f"Converted {ext} → Markdown ({lines} lines, {size} chars)")
    print(f"Output: {output_path}")


if __name__ == "__main__":
    main()
