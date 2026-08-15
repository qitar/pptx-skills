"""
utils.py - Shared utilities for pptx-compare.
"""

import difflib
import re
import xml.dom.minidom
from pathlib import Path


def parse_pretty(raw: bytes) -> str:
    try:
        parsed = xml.dom.minidom.parseString(raw)
        pretty = parsed.toprettyxml(indent="  ")
        lines = pretty.splitlines()
        if lines and lines[0].startswith("<?xml"):
            lines = lines[1:]
        return "\n".join(lines)
    except Exception:
        return raw.decode("utf-8", errors="replace")


def normalize_target(target: str) -> str:
    """Normalize a rels Target attribute to a 'ppt/...' path."""
    if target.startswith("/"):
        return target.lstrip("/")
    cleaned = re.sub(r"^(\.\./|\.\/)+", "", target)
    return "ppt/" + cleaned


def get_chart_map(slide_path: str, xml_contents: dict[str, str]) -> dict[str, str]:
    """Return {rId: 'ppt/charts/chartN.xml'} for a given slide by reading its .rels."""
    rels_path = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
    rels_xml = xml_contents.get(rels_path, "")
    if not rels_xml:
        return {}
    chart_map = {}
    try:
        dom = xml.dom.minidom.parseString(rels_xml.encode("utf-8"))
    except Exception:
        return {}
    chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    for rel in dom.getElementsByTagName("Relationship"):
        if rel.getAttribute("Type") == chart_type:
            rid = rel.getAttribute("Id")
            target = rel.getAttribute("Target")
            chart_map[rid] = normalize_target(target)
    return chart_map


def get_notes_path(slide_path: str, xml_contents: dict[str, str]) -> str | None:
    """Return 'ppt/notesSlides/notesSlideN.xml' for a given slide, or None if no notes."""
    rels_path = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
    rels_xml = xml_contents.get(rels_path, "")
    if not rels_xml:
        return None
    try:
        dom = xml.dom.minidom.parseString(rels_xml.encode("utf-8"))
    except Exception:
        return None
    notes_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    for rel in dom.getElementsByTagName("Relationship"):
        if rel.getAttribute("Type") == notes_type:
            return normalize_target(rel.getAttribute("Target"))
    return None


def write_diff(
    before_lines: list[str],
    after_lines: list[str],
    from_label: str,
    to_label: str,
    out_name: str,
    output_dir: Path,
) -> bool:
    diff = list(difflib.unified_diff(before_lines, after_lines,
                                     fromfile=from_label, tofile=to_label))
    if not diff:
        return False
    out_path = output_dir / out_name
    out_path.write_text("".join(diff), encoding="utf-8")
    return True
