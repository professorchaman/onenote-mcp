"""Convert markdown text to OneNote XML fragments.

Translates common markdown constructs into OneNote OE elements
using the xml_builder module. Handles headings, bold, italic,
lists, tables, and plain paragraphs.
"""

import re

from onenote_lib.xml_builder import (
    build_heading,
    build_list_item,
    build_outline,
    build_table,
    build_text_oe,
)


def markdown_to_outline(md: str) -> str:
    """Convert a markdown string into a OneNote Outline XML fragment.

    Supported markdown:
    - # Headings (levels 1-4)
    - **bold** and *italic*
    - - Bullet lists and 1. Numbered lists
    - | Pipe tables |
    - Plain paragraphs
    """
    lines = md.split("\n")
    oe_elements: list[str] = []
    i = 0

    while i < len(lines):
        line = lines[i]

        if not line.strip():
            i += 1
            continue

        heading_match = re.match(r"^(#{1,4})\s+(.+)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            oe_elements.append(build_heading(heading_match.group(2).strip(), level))
            i += 1
            continue

        if line.lstrip().startswith("|") and line.rstrip().endswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                table_lines.append(lines[i])
                i += 1
            oe_elements.append(_parse_md_table(table_lines))
            continue

        bullet_match = re.match(r"^(\s*)[-*+]\s+(.+)$", line)
        if bullet_match:
            oe_elements.append(build_list_item(bullet_match.group(2).strip()))
            i += 1
            continue

        num_match = re.match(r"^(\s*)\d+[.)]\s+(.+)$", line)
        if num_match:
            oe_elements.append(build_list_item(num_match.group(2).strip(), numbered=True))
            i += 1
            continue

        bold, italic, text = _detect_formatting(line)
        oe_elements.append(build_text_oe(text, bold=bold, italic=italic))
        i += 1

    return build_outline(oe_elements)


def _parse_md_table(lines: list[str]) -> str:
    """Parse markdown table lines into a build_table call."""
    rows: list[list[str]] = []
    has_header = False
    for line in lines:
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.match(r"^:?-+:?$", c) for c in cells):
            has_header = True
            continue
        rows.append(cells)
    return build_table(rows, has_header=has_header)


def _detect_formatting(text: str) -> tuple[bool, bool, str]:
    """Detect if the entire line is bold/italic and strip markers.

    Returns (is_bold, is_italic, cleaned_text).
    Only detects whole-line formatting — inline mixed formatting
    is passed through as-is for now.
    """
    stripped = text.strip()
    if stripped.startswith("***") and stripped.endswith("***"):
        return True, True, stripped[3:-3]
    if stripped.startswith("**") and stripped.endswith("**"):
        return True, False, stripped[2:-2]
    if stripped.startswith("*") and stripped.endswith("*") and len(stripped) > 2:
        return False, True, stripped[1:-1]
    return False, False, stripped
