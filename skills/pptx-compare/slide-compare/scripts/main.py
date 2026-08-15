"""
main.py - Slide Compare: full XML diff with minimal normalization.

Uses minimal normalization + node-based diff.
For small XML files (< 200 lines), outputs both before/after for direct reading.

Usage:
    PYTHONPATH=skills/pptx-compare \
      python skills/pptx-compare/slide-compare/scripts/main.py \
      before.pptx after.pptx
"""

import argparse
import json
import sys
import tempfile
from pathlib import Path

from core.differ import generate_diffs
from core.extractor import extract_master_data, extract_slides
from core.mapper import match_by_filename, match_slides


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Slide Compare: full XML diff with minimal normalization."
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

    # Extract slides with minimal normalization for diff
    before_slides, _ = extract_slides(before_path, apply_normalization=False)
    after_slides,  _ = extract_slides(after_path,  apply_normalization=False)

    # Re-extract with heavy normalization for reliable Jaccard mapping
    before_slides_norm, _ = extract_slides(before_path, apply_normalization=True)
    after_slides_norm,  _ = extract_slides(after_path,  apply_normalization=True)

    # Extract masters/layouts/themes with minimal normalization for diff
    before_masters = extract_master_data(before_path, apply_normalization=False)
    after_masters  = extract_master_data(after_path,  apply_normalization=False)

    # Extract masters with heavy normalization for reliable mapping
    before_masters_norm = extract_master_data(before_path, apply_normalization=True)
    after_masters_norm  = extract_master_data(after_path,  apply_normalization=True)

    # Map slides via Jaccard (no CRC pre-pass — always go to diff)
    matches = match_slides(before_slides_norm, after_slides_norm, threshold=args.threshold)

    # Map masters by filename (no CRC pre-pass — always go to diff)
    master_matches = match_by_filename(before_masters_norm, after_masters_norm)

    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        output_dir = Path(tempfile.mkdtemp(prefix="pptx_full_compare_"))

    # Generate node-based diffs for slides
    generate_diffs(
        before_slides=before_slides,
        after_slides=after_slides,
        matches=matches,
        before_label=Path(before_path).name,
        after_label=Path(after_path).name,
        output_dir=output_dir,
        mode="minimal",
    )

    # Generate node-based diffs for masters/layouts/themes
    generate_diffs(
        before_slides=before_masters,
        after_slides=after_masters,
        matches=master_matches,
        before_label=Path(before_path).name,
        after_label=Path(after_path).name,
        output_dir=output_dir / "masters",
        mode="minimal",
    )

    # Build output
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
        mapping_entries.append(entry)

    before_masters_by_idx = {s["index"]: s for s in before_masters_norm}
    after_masters_by_idx  = {s["index"]: s for s in after_masters_norm}
    master_mapping_entries = []
    for m in master_matches:
        entry = {
            "filename":     (before_masters_by_idx.get(m.before_index) or after_masters_by_idx.get(m.after_index))["filename"],  # type: ignore[index]
            "before_index": m.before_index,
            "after_index":  m.after_index,
            "status":       m.status,
        }
        if m.diff_file:
            entry["diff_file"] = m.diff_file
        master_mapping_entries.append(entry)

    output = {
        "mapping":        mapping_entries,
        "master_mapping": master_mapping_entries,
        "diff_dir":       str(output_dir),
    }

    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
