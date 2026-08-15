"""
main.py - Slide Compare Lite: normalized XML diff for PPTX comparison.

Uses heavy PowerPoint-aware normalization + unified diff.
Outputs file metadata, summary, and per-slide diffs in a single pass.

Usage:
    PYTHONPATH=skills/pptx-compare \
      python skills/pptx-compare/slide-compare-lite/scripts/main.py \
      before.pptx after.pptx
"""

import argparse
import json
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from core.differ import generate_diffs
from core.extractor import extract_checksums, extract_master_data, extract_slides
from core.mapper import match_by_filename, match_slides


def _format_mtime(path: str) -> str:
    return datetime.fromtimestamp(os.path.getmtime(path)).astimezone().isoformat(timespec="seconds")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Slide Compare Lite: normalized XML diff with PowerPoint-noise normalization."
    )
    parser.add_argument("before_pptx", help="Path to the before PPTX file")
    parser.add_argument("after_pptx",  help="Path to the after PPTX file")
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.5,
        metavar="FLOAT",
        help="Similarity threshold for slide matching (default: 0.5)",
    )
    parser.add_argument(
        "--output-dir",
        metavar="DIR",
        help="Directory to write diff files. Defaults to a temporary directory.",
    )
    args = parser.parse_args()

    before_path = args.before_pptx
    after_path  = args.after_pptx

    if not Path(before_path).exists():
        print(json.dumps({"error": f"File not found: {before_path}"}))
        sys.exit(1)
    if not Path(after_path).exists():
        print(json.dumps({"error": f"File not found: {after_path}"}))
        sys.exit(1)

    # File metadata (pure I/O, no XML processing)
    size_before  = os.path.getsize(before_path)
    size_after   = os.path.getsize(after_path)
    mtime_before = _format_mtime(before_path)
    mtime_after  = _format_mtime(after_path)

    # Pass 1: checksum pre-match (ZIP header only, no XML parsing)
    before_checksums = extract_checksums(before_path)
    after_checksums  = extract_checksums(after_path)

    # Extract with strong normalization
    before_slides, _ = extract_slides(before_path, apply_normalization=True)
    after_slides,  _ = extract_slides(after_path,  apply_normalization=True)

    # Map slides: checksum identical → skip Jaccard, remainder → Jaccard
    matches = match_slides(
        before_slides, after_slides,
        threshold=args.threshold,
        before_checksums=before_checksums,
        after_checksums=after_checksums,
    )

    # Extract and match masters / layouts / themes
    before_masters = extract_master_data(before_path)
    after_masters  = extract_master_data(after_path)
    master_matches = match_by_filename(
        before_masters, after_masters,
        before_checksums=before_checksums,
        after_checksums=after_checksums,
    )

    # Generate diffs
    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        output_dir = Path(tempfile.mkdtemp(prefix="pptx_lite_compare_"))

    generate_diffs(
        before_slides=before_slides,
        after_slides=after_slides,
        matches=matches,
        before_label=Path(before_path).name,
        after_label=Path(after_path).name,
        output_dir=output_dir,
        mode="lite",
    )

    generate_diffs(
        before_slides=before_masters,
        after_slides=after_masters,
        matches=master_matches,
        before_label=Path(before_path).name,
        after_label=Path(after_path).name,
        output_dir=output_dir / "masters",
        mode="lite",
    )

    before_by_idx = {s["index"]: s for s in before_slides}
    after_by_idx  = {s["index"]: s for s in after_slides}
    before_masters_by_idx = {s["index"]: s for s in before_masters}
    after_masters_by_idx  = {s["index"]: s for s in after_masters}

    # Build master mapping entries
    master_mapping_entries = []
    for m in master_matches:
        entry: dict = {
            "filename":    (before_masters_by_idx.get(m.before_index) or after_masters_by_idx.get(m.after_index))["filename"],  # type: ignore[index]
            "before_index": m.before_index,
            "after_index":  m.after_index,
            "status":       m.status,
            "similarity":   m.similarity,
        }
        if m.diff_file:
            entry["diff_file"] = m.diff_file
        if m.before_index is not None and m.before_index in before_masters_by_idx:
            entry["before_checksum"] = before_masters_by_idx[m.before_index]["checksum_hex"]
        if m.after_index is not None and m.after_index in after_masters_by_idx:
            entry["after_checksum"] = after_masters_by_idx[m.after_index]["checksum_hex"]
        master_mapping_entries.append(entry)

    # Build mapping entries
    mapping_entries = []
    for m in matches:
        entry: dict = {
            "before_index": m.before_index,
            "after_index":  m.after_index,
            "status":       m.status,
            "similarity":   m.similarity,
        }
        if m.diff_file:
            entry["diff_file"] = m.diff_file
        if m.before_index is not None and m.before_index in before_by_idx:
            b = before_by_idx[m.before_index]
            entry["before_creation_id"] = b["creation_id"]
            entry["before_checksum"]    = b["checksum_hex"]
        if m.after_index is not None and m.after_index in after_by_idx:
            a = after_by_idx[m.after_index]
            entry["after_creation_id"] = a["creation_id"]
            entry["after_checksum"]    = a["checksum_hex"]
        mapping_entries.append(entry)

    output = {
        "master_summary": {
            "identical": sum(1 for m in master_matches if m.status == "identical"),
            "modified":  sum(1 for m in master_matches if m.status == "modified"),
            "added":     sum(1 for m in master_matches if m.status == "added"),
            "deleted":   sum(1 for m in master_matches if m.status == "deleted"),
        },
        "master_mapping": master_mapping_entries,
        "before_meta": {
            "path":        before_path,
            "size_bytes":  size_before,
            "modified_at": mtime_before,
            "slide_count": len(before_slides),
        },
        "after_meta": {
            "path":        after_path,
            "size_bytes":  size_after,
            "modified_at": mtime_after,
            "slide_count": len(after_slides),
        },
        "summary": {
            "identical": sum(1 for m in matches if m.status == "identical"),
            "modified":  sum(1 for m in matches if m.status == "modified"),
            "added":     sum(1 for m in matches if m.status == "added"),
            "deleted":   sum(1 for m in matches if m.status == "deleted"),
        },
        "diff_dir": str(output_dir),
        "mapping":  mapping_entries,
    }

    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
