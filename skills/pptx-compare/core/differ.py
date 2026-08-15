"""
differ.py - Generate diffs for matched slide pairs.

Supports two modes:
- "lite":    Heavy PowerPoint-noise normalization + unified diff (default for slide-compare-lite)
- "minimal": Minimal normalization + node-based diff (for slide-compare)

For small XML files (below direct_threshold lines), outputs raw normalized XML
for both before and after so Claude can read them directly.
"""

from pathlib import Path
from typing import Literal

from core.extractor import SlideData
from core.mapper import SlideMatch
from core.normalizer import (
    to_node_lines_minimal,
    to_normalized_lines_lite,
)
from core.utils import write_diff

# When normalized XML is shorter than this, write both sides for direct reading
DIRECT_READ_THRESHOLD = 200  # lines


def generate_diffs(
    before_slides: list[SlideData],
    after_slides: list[SlideData],
    matches: list[SlideMatch],
    before_label: str,
    after_label: str,
    output_dir: Path,
    mode: Literal["lite", "minimal"] = "lite",
) -> None:
    """
    Generate diff files for all 'modified' slide pairs.

    - mode="lite":    Uses heavy normalization + unified diff.
    - mode="minimal": Uses minimal normalization + node-based diff.
                      For short XMLs, writes raw XML for direct reading instead.

    Updates each SlideMatch.diff_file in place.
    Sets status to "identical" if diff is empty after normalization.
    """
    before_by_index = {s["index"]: s for s in before_slides}
    after_by_index  = {s["index"]: s for s in after_slides}

    output_dir.mkdir(parents=True, exist_ok=True)

    for match in matches:
        if match.status != "modified":
            continue

        b = before_by_index[match.before_index]
        a = after_by_index[match.after_index]

        out_stem = f"slide_B{match.before_index}_A{match.after_index}"

        if mode == "lite":
            b_lines = to_normalized_lines_lite(b["xml_content"])
            a_lines = to_normalized_lines_lite(a["xml_content"])
            out_name = out_stem + ".diff"
            wrote = write_diff(
                b_lines, a_lines,
                f"{before_label}/{b['filename']}",
                f"{after_label}/{a['filename']}",
                out_name,
                output_dir,
            )
        else:
            # minimal mode: node-based diff
            b_lines = to_node_lines_minimal(b["xml_content"])
            a_lines = to_node_lines_minimal(a["xml_content"])

            if len(b_lines) < DIRECT_READ_THRESHOLD and len(a_lines) < DIRECT_READ_THRESHOLD:
                # Small XML: write both sides for Claude to read directly.
                # Compare first — writing a file whose two halves are identical
                # would report an unchanged slide as "modified".
                out_name = out_stem + "_direct.txt"
                wrote = b_lines != a_lines
                if wrote:
                    _write_direct(b_lines, a_lines, out_stem, output_dir,
                                  before_label, b["filename"],
                                  after_label, a["filename"])
            else:
                out_name = out_stem + ".diff"
                wrote = write_diff(
                    b_lines, a_lines,
                    f"{before_label}/{b['filename']}",
                    f"{after_label}/{a['filename']}",
                    out_name,
                    output_dir,
                )

        if wrote:
            match.diff_file = str(output_dir / out_name)
        else:
            match.status = "identical"


def _write_direct(
    b_lines: list[str],
    a_lines: list[str],
    out_stem: str,
    output_dir: Path,
    before_label: str,
    before_filename: str,
    after_label: str,
    after_filename: str,
) -> None:
    """Write before and after XML side-by-side for direct Claude reading."""
    out_path = output_dir / (out_stem + "_direct.txt")
    content = (
        f"=== BEFORE: {before_label}/{before_filename} ===\n"
        + "".join(b_lines)
        + f"\n=== AFTER: {after_label}/{after_filename} ===\n"
        + "".join(a_lines)
    )
    out_path.write_text(content, encoding="utf-8")
