"""
mapper.py - Match slides between two PPTX files using Jaccard similarity.

Algorithm:
1. Jaccard similarity scoring for all slide pairs (namespace/URL tokens excluded)
2. Position bias: score = 0.8 * jaccard + 0.2 * position_similarity
3. Greedy 1:1 assignment (highest score first, threshold applied on raw Jaccard)
4. Pairs below threshold → added / deleted
"""

import re
from dataclasses import dataclass, field
from typing import Literal

from core.extractor import SlideData

_NOISE_RE = re.compile(r'xmlns:|http://|https://')


def _tokenize(xml_content: str) -> set[str]:
    return {t for t in xml_content.split() if not _NOISE_RE.search(t)}


@dataclass
class SlideMatch:
    before_index: int | None  # 1-based; None = added slide
    after_index: int | None   # 1-based; None = deleted slide
    similarity: float            # raw Jaccard similarity
    status: Literal["identical", "modified", "added", "deleted"]
    diff_file: str | None = field(default=None)


def _checksum_identical(b_crc: str | None, a_crc: str | None) -> bool:
    """True only when both checksums are present and equal."""
    return b_crc is not None and a_crc is not None and b_crc == a_crc


def match_by_filename(
    before_items: list[SlideData],
    after_items: list[SlideData],
    before_checksums: dict[str, str] | None = None,
    after_checksums: dict[str, str] | None = None,
) -> list[SlideMatch]:
    """
    Match items by filename (for masters, layouts, themes).
    No Jaccard — filename equality determines pairing.
    Checksum equality → identical immediately.
    """
    results: list[SlideMatch] = []
    after_by_filename = {s["filename"]: s for s in after_items}
    matched_after: set[str] = set()

    for b in before_items:
        if b["filename"] not in after_by_filename:
            results.append(SlideMatch(
                before_index=b["index"],
                after_index=None,
                similarity=0.0,
                status="deleted",
            ))
            continue

        a = after_by_filename[b["filename"]]
        matched_after.add(b["filename"])

        if before_checksums and after_checksums and _checksum_identical(
            before_checksums.get(b["filename"]), after_checksums.get(a["filename"])
        ):
            results.append(SlideMatch(
                before_index=b["index"],
                after_index=a["index"],
                similarity=1.0,
                status="identical",
            ))
            continue

        results.append(SlideMatch(
            before_index=b["index"],
            after_index=a["index"],
            similarity=0.0,
            status="modified",
        ))

    for a in after_items:
        if a["filename"] not in matched_after:
            results.append(SlideMatch(
                before_index=None,
                after_index=a["index"],
                similarity=0.0,
                status="added",
            ))

    return results


def _checksum_prematch(
    before_slides: list[SlideData],
    after_slides: list[SlideData],
    before_checksums: dict[str, str] | None,
    after_checksums: dict[str, str] | None,
) -> tuple[list[SlideMatch], set[int], set[int]]:
    """
    Pass 1: pre-match slides identical by filename + CRC32 (no XML parsing).

    Returns (identical_matches, used_before_indices, used_after_indices).
    A no-op when either checksum map is missing.
    """
    results: list[SlideMatch] = []
    used_before: set[int] = set()
    used_after: set[int] = set()
    if not (before_checksums and after_checksums):
        return results, used_before, used_after

    before_by_filename = {s["filename"]: s for s in before_slides}
    after_by_filename  = {s["filename"]: s for s in after_slides}
    for filename, b in before_by_filename.items():
        a = after_by_filename.get(filename)
        if a is None:
            continue
        if _checksum_identical(before_checksums.get(filename), after_checksums.get(filename)):
            results.append(SlideMatch(
                before_index=b["index"],
                after_index=a["index"],
                similarity=1.0,
                status="identical",
            ))
            used_before.add(b["index"])
            used_after.add(a["index"])
    return results, used_before, used_after


def _score_pairs(
    remaining_before: list[SlideData],
    remaining_after: list[SlideData],
    nb: int,
    na: int,
) -> list[tuple[float, float, int, int]]:
    """
    Score every remaining before/after pair as (score, jaccard, b_idx, a_idx).

    score = 0.8 * jaccard + 0.2 * position_similarity. Position is normalized
    against the original total counts (nb, na) to preserve relative positions.
    """
    before_tokens = {s["index"]: _tokenize(s["xml_content"]) for s in remaining_before}
    after_tokens  = {s["index"]: _tokenize(s["xml_content"]) for s in remaining_after}

    scores: list[tuple[float, float, int, int]] = []
    for b in remaining_before:
        tb = before_tokens[b["index"]]
        b_norm = (b["index"] - 1) / max(nb - 1, 1)
        for a in remaining_after:
            ta = after_tokens[a["index"]]
            intersection = len(tb & ta)
            union = len(tb | ta)
            jaccard = intersection / union if union else 0.0
            a_norm = (a["index"] - 1) / max(na - 1, 1)
            pos_sim = 1.0 - abs(b_norm - a_norm)
            score = 0.8 * jaccard + 0.2 * pos_sim
            scores.append((score, jaccard, b["index"], a["index"]))
    return scores


def _greedy_assign(
    scores: list[tuple[float, float, int, int]],
    threshold: float,
    used_before: set[int],
    used_after: set[int],
) -> list[tuple[float, int, int]]:
    """
    Greedy 1:1 assignment over scores (highest score first).

    Stops once the best remaining raw Jaccard drops below `threshold`.
    Mutates used_before/used_after; returns [(jaccard, b_idx, a_idx)].
    """
    greedy_matches: list[tuple[float, int, int]] = []
    for _score, jaccard, b_idx, a_idx in sorted(scores, key=lambda x: x[0], reverse=True):
        if b_idx in used_before or a_idx in used_after:
            continue
        if jaccard < threshold:
            break
        greedy_matches.append((jaccard, b_idx, a_idx))
        used_before.add(b_idx)
        used_after.add(a_idx)
    return greedy_matches


def _unmatched_matches(
    remaining: list[SlideData],
    used: set[int],
    *,
    side: Literal["before", "after"],
) -> list[SlideMatch]:
    """Build deleted (side='before') / added (side='after') matches for leftovers."""
    results: list[SlideMatch] = []
    for s in remaining:
        if s["index"] in used:
            continue
        if side == "before":
            results.append(SlideMatch(s["index"], None, 0.0, "deleted"))
        else:
            results.append(SlideMatch(None, s["index"], 0.0, "added"))
    return results


def match_slides(
    before_slides: list[SlideData],
    after_slides: list[SlideData],
    threshold: float = 0.5,
    before_checksums: dict[str, str] | None = None,
    after_checksums: dict[str, str] | None = None,
) -> list[SlideMatch]:
    """
    Match before slides to after slides and return a mapping list.

    Pass 1 (optional): if checksums are provided, pre-match slides that share the
    same filename and CRC32 — these are confirmed identical without XML parsing.

    Pass 2: Jaccard similarity on remaining unmatched slides, with greedy 1:1
    assignment and position bias (0.8 * jaccard + 0.2 * position).
    Pairs below `threshold` are treated as added/deleted.
    """
    results, used_before, used_after = _checksum_prematch(
        before_slides, after_slides, before_checksums, after_checksums
    )

    # Pass 2: Jaccard on unmatched slides
    remaining_before = [s for s in before_slides if s["index"] not in used_before]
    remaining_after  = [s for s in after_slides  if s["index"] not in used_after]

    # Use original total counts for position normalization to preserve relative positions
    scores = _score_pairs(remaining_before, remaining_after, len(before_slides), len(after_slides))
    greedy_matches = _greedy_assign(scores, threshold, used_before, used_after)

    for jaccard, b_idx, a_idx in greedy_matches:
        results.append(SlideMatch(
            before_index=b_idx,
            after_index=a_idx,
            similarity=round(jaccard, 6),
            status="modified",
        ))

    results.extend(_unmatched_matches(remaining_before, used_before, side="before"))
    results.extend(_unmatched_matches(remaining_after, used_after, side="after"))

    results.sort(key=lambda m: (
        m.before_index if m.before_index is not None else float("inf"),
        m.after_index  if m.after_index  is not None else float("inf"),
    ))

    return results
