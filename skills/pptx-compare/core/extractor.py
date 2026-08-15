"""
extractor.py - Extract slide XML data from a PPTX file.
"""

import posixpath
import re
import xml.dom.minidom
import zipfile
from typing import TypedDict

from core.normalizer import structural_normalize
from core.utils import get_chart_map, get_notes_path, normalize_target, parse_pretty
from lxml import etree as _lxml_etree  # type: ignore[import-untyped]

_NOISY_RE = (
    re.compile(r'<p14:modId\b[^/]*/>', re.DOTALL),
    re.compile(r'<p14:creationId\b[^/]*/>', re.DOTALL),
    re.compile(r'\s+dirty="[^"]*"'),
    re.compile(r'\s+err="[^"]*"'),
)

_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _strip_noisy(xml: str) -> str:
    for pat in _NOISY_RE:
        xml = pat.sub('', xml)
    return xml


def _embed_charts(xml_content: str, chart_map: dict[str, str], all_xml: dict[str, str]) -> str:
    """Replace <c:chart r:id="rIdX"/> references with the actual chart XML content."""
    for rid, chart_path in chart_map.items():
        chart_xml = all_xml.get(chart_path, "")
        if not chart_xml:
            continue
        chart_xml = _strip_noisy(chart_xml)
        xml_content = re.sub(
            rf'<c:chart\b[^>]*\br:id="{re.escape(rid)}"[^/]*/>', chart_xml, xml_content
        )
    return xml_content


def _embed_notes(xml_content: str, slide_path: str, all_xml: dict[str, str]) -> str:
    """Append notes slide XML as the last child of the slide root element."""
    notes_path = get_notes_path(slide_path, all_xml)
    if not notes_path or notes_path not in all_xml:
        return xml_content
    notes_xml = _strip_noisy(all_xml[notes_path])
    try:
        slide_root = _lxml_etree.fromstring(xml_content.encode("utf-8"))
        notes_root = _lxml_etree.fromstring(notes_xml.encode("utf-8"))
        slide_root.append(notes_root)
        return _lxml_etree.tostring(slide_root, encoding="unicode")
    except _lxml_etree.XMLSyntaxError:
        return xml_content


def _extract_p14_ids(raw: bytes) -> tuple[str | None, str | None]:
    """Extract (creation_id, mod_id) from raw slide XML bytes, or None for each absent."""
    raw_str = raw.decode("utf-8", errors="replace")
    creation_match = re.search(r'p14:creationId\b[^>]*/>', raw_str)
    mod_match = re.search(r'p14:modId\b[^>]*/>', raw_str)
    creation_id = re.search(r'val="(\d+)"', creation_match.group()).group(1) if creation_match else None
    mod_id = re.search(r'val="(\d+)"', mod_match.group()).group(1) if mod_match else None
    return creation_id, mod_id


class SlideData(TypedDict):
    index: int                  # 1-based slide order
    filename: str               # "ppt/slides/slideN.xml"
    xml_content: str            # slide XML with chart XMLs and notes XML embedded inline
    chart_map: dict[str, str]   # {rId: "ppt/charts/chartN.xml"}
    creation_id: str | None     # p14:creationId val
    mod_id: str | None          # p14:modId val
    checksum: int             # CRC32 from ZIP entry header for slide XML
    checksum_hex: str         # Hex string representation (e.g., 0x1a2b3c4d)


def _make_slide_data(
    index: int,
    filename: str,
    pre_norm_xml: str,
    *,
    chart_map: dict[str, str],
    creation_id: str | None,
    mod_id: str | None,
    crc_map: dict[str, int],
    apply_normalization: bool,
) -> SlideData:
    """
    Assemble a SlideData from already strip/embed-processed XML.

    Applies structural_normalize only when requested, then attaches the CRC32
    checksum. Shared by extract_slides and extract_master_data.
    """
    xml_content = structural_normalize(pre_norm_xml) if apply_normalization else pre_norm_xml
    checksum = crc_map.get(filename, 0)
    return SlideData(
        index=index,
        filename=filename,
        xml_content=xml_content,
        chart_map=chart_map,
        creation_id=creation_id,
        mod_id=mod_id,
        checksum=checksum,
        checksum_hex=f"0x{checksum:08x}",
    )


def extract_slides(
    pptx_path: str,
    apply_normalization: bool = True,
) -> tuple[list[SlideData], dict[str, str]]:
    """
    Extract slide data from a PPTX file.

    Args:
        pptx_path: Path to PPTX file.
        apply_normalization: If True, apply structural_normalize() (for slide-compare-lite,
                             and for slide-compare's mapping pass).
                             If False, only strip noisy regex + embed charts (for slide-compare).

    Returns:
        (slides, all_xml) where:
        - slides: list of SlideData in presentation order
        - all_xml: dict of all XML/rels file contents keyed by path
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        rid_to_file: dict[str, str] = {}
        zip_names = zf.namelist()
        crc_map = {info.filename: info.CRC for info in zf.infolist()}
        with zf.open("ppt/_rels/presentation.xml.rels") as f:
            dom = xml.dom.minidom.parse(f)
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if "slides/slide" in target and "slideLayout" not in target and "slideMaster" not in target:
                rid_to_file[rel.getAttribute("Id")] = normalize_target(target)

        ordered_filenames: list[str] = []
        with zf.open("ppt/presentation.xml") as f:
            dom = xml.dom.minidom.parse(f)
        for sld in dom.getElementsByTagNameNS("*", "sldId"):
            rid = sld.getAttribute("r:id")
            if rid in rid_to_file:
                ordered_filenames.append(rid_to_file[rid])

        ordered_set = set(ordered_filenames)
        for name in zip_names:
            if (
                name.startswith("ppt/slides/slide")
                and name.endswith(".xml")
                and name not in ordered_set
            ):
                ordered_filenames.append(name)
                ordered_set.add(name)

        all_xml: dict[str, str] = {}
        raw_map: dict[str, bytes] = {}
        for name in zip_names:
            if name.endswith(".xml") or name.endswith(".rels"):
                raw = zf.read(name)
                raw_map[name] = raw
                all_xml[name] = parse_pretty(raw)

    slides: list[SlideData] = []
    for i, filename in enumerate(ordered_filenames, start=1):
        raw = raw_map.get(filename, b"")
        chart_map = get_chart_map(filename, all_xml)

        xml_content = _strip_noisy(all_xml.get(filename, ""))
        xml_content = _embed_charts(xml_content, chart_map, all_xml)
        xml_content = _embed_notes(xml_content, filename, all_xml)

        creation_id, mod_id = _extract_p14_ids(raw)

        slides.append(_make_slide_data(
            i, filename, xml_content,
            chart_map=chart_map,
            creation_id=creation_id,
            mod_id=mod_id,
            crc_map=crc_map,
            apply_normalization=apply_normalization,
        ))

    return slides, all_xml


# Relationship types that should NOT be included as separate analysis targets.
# Slides are handled by extract_slides; charts/notes are embedded inline into slides;
# images/media/embeddings are binary; package-level files are not slide content.
_NON_CONTENT_REL_TYPES: frozenset[str] = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/thumbnail",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
})


def _collect_non_slide_xml_targets(
    zip_names: list[str],
    raw_map: dict[str, bytes],
) -> list[str]:
    """
    Traverse all .rels files in the ZIP and collect XML targets that are not
    slides, charts, images, or other non-content types.
    Path resolution uses posixpath so ../ segments are handled correctly.
    """
    zip_name_set = set(zip_names)
    found: set[str] = set()

    for name in zip_names:
        if not name.endswith(".rels"):
            continue
        # Charts are embedded inline into slides — skip their sub-rels entirely
        if "ppt/charts/" in name:
            continue
        raw = raw_map.get(name, b"")
        if not raw:
            continue
        try:
            dom = xml.dom.minidom.parseString(raw)
        except Exception:
            continue
        for rel in dom.getElementsByTagName("Relationship"):
            if rel.getAttribute("Type") in _NON_CONTENT_REL_TYPES:
                continue
            target = rel.getAttribute("Target")
            if not target.endswith(".xml"):
                continue
            # e.g. "ppt/slides/_rels/slide1.xml.rels" → base dir is "ppt/slides/"
            actual_dir = name.replace("/_rels/", "/").rsplit("/", 1)[0]
            resolved = posixpath.normpath(actual_dir + "/" + target).lstrip("/")
            if resolved in zip_name_set and resolved.startswith("ppt/"):
                found.add(resolved)

    return sorted(found)


def extract_master_data(
    pptx_path: str,
    apply_normalization: bool = True,
) -> list[SlideData]:
    """
    Extract non-slide XML content (masters, layouts, themes, etc.) from a PPTX file.
    Targets are discovered by traversing .rels files — no hardcoded filename patterns.
    Returns SlideData list sorted by filename. No chart/notes embedding.
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zip_names = zf.namelist()
        crc_map = {info.filename: info.CRC for info in zf.infolist()}
        all_raw: dict[str, bytes] = {
            n: zf.read(n) for n in zip_names
            if n.endswith(".xml") or n.endswith(".rels")
        }

    target_names = _collect_non_slide_xml_targets(zip_names, all_raw)
    raw_map = {n: all_raw[n] for n in target_names if n in all_raw}

    items: list[SlideData] = []
    for i, filename in enumerate(target_names, start=1):
        xml_content = _strip_noisy(parse_pretty(raw_map[filename]))
        items.append(_make_slide_data(
            i, filename, xml_content,
            chart_map={},
            creation_id=None,
            mod_id=None,
            crc_map=crc_map,
            apply_normalization=apply_normalization,
        ))

    return items


def extract_checksums(pptx_path: str) -> dict[str, str]:
    """
    Return CRC32 checksums for all analysis targets (slides + rels-discovered XML),
    keyed by filename.
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zip_names = zf.namelist()
        crc_map = {info.filename: info.CRC for info in zf.infolist()}
        rels_raw: dict[str, bytes] = {
            n: zf.read(n) for n in zip_names if n.endswith(".rels")
        }

    slide_targets = {
        n for n in zip_names
        if n.startswith("ppt/slides/slide") and n.endswith(".xml")
    }
    non_slide_targets = set(_collect_non_slide_xml_targets(zip_names, rels_raw))
    all_targets = slide_targets | non_slide_targets

    return {
        name: f"0x{crc_map[name]:08x}"
        for name in all_targets
        if name in crc_map
    }
