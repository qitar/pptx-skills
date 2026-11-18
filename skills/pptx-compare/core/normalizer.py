"""
normalizer.py - XML normalization rules for PPTX comparison.

Two normalization levels:
- structural_normalize(): Strong normalization for Jaccard mapping prep
  Removes PowerPoint-generated structural noise (used by slide-compare-lite)
- normalize_element_lite(): Additional normalization for diff generation (lite mode)
- normalize_element_minimal(): Minimal normalization for diff generation (full mode)
"""

from lxml import etree  # type: ignore[import-untyped]

_A_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"
_C_NS   = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

_P14_NOISY_TAGS = {f"{{{_P14_NS}}}modId", f"{{{_P14_NS}}}creationId"}

_C_V_TAG            = f"{{{_C_NS}}}v"
_C_SCALING_TAG      = f"{{{_C_NS}}}scaling"
_C_LANG_TAG         = f"{{{_C_NS}}}lang"
_C_SPR_TAG          = f"{{{_C_NS}}}spPr"
_C_SER_TAG          = f"{{{_C_NS}}}ser"
_C_ORDER_TAG        = f"{{{_C_NS}}}order"
_C_CHART_TAG        = f"{{{_C_NS}}}chart"
_C_EXT_DATA_TAG     = f"{{{_C_NS}}}externalData"
_C_AUTO_UPD_TAG     = f"{{{_C_NS}}}autoUpdate"
_A_SOLID_FILL       = f"{{{_A_NS}}}solidFill"
_A_DEF_RPR_TAG      = f"{{{_A_NS}}}defRPr"
_A_END_PARA_RPR_TAG = f"{{{_A_NS}}}endParaRPr"
_A_EXT_TAG          = f"{{{_A_NS}}}ext"
_A_EXT_LST_TAG      = f"{{{_A_NS}}}extLst"
_A_R_TAG            = f"{{{_A_NS}}}r"
_A_T_TAG            = f"{{{_A_NS}}}t"
_A_P_TAG            = f"{{{_A_NS}}}p"
_A_TBL_TAG          = f"{{{_A_NS}}}tbl"
_A_TBL_GRID_TAG     = f"{{{_A_NS}}}tblGrid"
_A_GRID_COL_TAG     = f"{{{_A_NS}}}gridCol"
_A_TC_TAG           = f"{{{_A_NS}}}tc"
_A_TC_PR_TAG        = f"{{{_A_NS}}}tcPr"
_A_FLD_TAG          = f"{{{_A_NS}}}fld"
_A_RPR_TAG          = f"{{{_A_NS}}}rPr"
_A_SCHEME_CLR_TAG   = f"{{{_A_NS}}}schemeClr"
_A_TABLE_BORDER_TAGS = {
    f"{{{_A_NS}}}lnT", f"{{{_A_NS}}}lnB",
    f"{{{_A_NS}}}lnL", f"{{{_A_NS}}}lnR",
}

_C_PT_TAG           = f"{{{_C_NS}}}pt"
_C_DPT_TAG          = f"{{{_C_NS}}}dPt"
_C_IDX_TAG          = f"{{{_C_NS}}}idx"
_C_AXIS_FLOAT_TAGS  = {
    f"{{{_C_NS}}}min", f"{{{_C_NS}}}max",
    f"{{{_C_NS}}}majorUnit", f"{{{_C_NS}}}minorUnit",
}
_C_NUM_CACHE_TAG    = f"{{{_C_NS}}}numCache"
_C_FORMAT_CODE_TAG  = f"{{{_C_NS}}}formatCode"
_C_NUM_FMT_TAG      = f"{{{_C_NS}}}numFmt"
_C_D_LBLS_TAG       = f"{{{_C_NS}}}dLbls"
_C_TX_PR_TAG        = f"{{{_C_NS}}}txPr"
_C_LBL_ALGN_TAG     = f"{{{_C_NS}}}lblAlgn"
_C_NO_MULTI_LVL_TAG = f"{{{_C_NS}}}noMultiLvlLbl"
_C_CROSSES_TAG      = f"{{{_C_NS}}}crosses"

_C16_NS             = "http://schemas.microsoft.com/office/drawing/2014/chart"
_C16_UNIQUE_ID_TAG  = f"{{{_C16_NS}}}uniqueId"

_R_NS               = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CHART_TYPE_CONTAINERS = {
    f"{{{_C_NS}}}barChart",    f"{{{_C_NS}}}lineChart",   f"{{{_C_NS}}}areaChart",
    f"{{{_C_NS}}}pieChart",    f"{{{_C_NS}}}doughnutChart", f"{{{_C_NS}}}radarChart",
    f"{{{_C_NS}}}scatterChart", f"{{{_C_NS}}}bubbleChart", f"{{{_C_NS}}}stockChart",
    f"{{{_C_NS}}}ofPieChart",  f"{{{_C_NS}}}bar3DChart",  f"{{{_C_NS}}}line3DChart",
    f"{{{_C_NS}}}area3DChart", f"{{{_C_NS}}}pie3DChart",  f"{{{_C_NS}}}surface3DChart",
    f"{{{_C_NS}}}surfaceChart",
}


_BLANK_TEXT_PARSER = etree.XMLParser(remove_blank_text=True)

_A_SRGB_CLR_TAG = f"{{{_A_NS}}}srgbClr"


# ---------------------------------------------------------------------------
# Shared element helpers
# ---------------------------------------------------------------------------

def _detach(el: etree._Element) -> None:
    """Remove an element from its parent, if it has one."""
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def _remove_all(root: etree._Element, tag: str) -> None:
    """Detach every descendant element whose tag matches `tag`."""
    for el in list(root.findall(f".//{tag}")):
        _detach(el)


def _normalize_float_str(s: str) -> str:
    """
    Canonicalize a numeric string to a stable 14-significant-figure form.
    PowerPoint stores the same float in different representations across saves
    (e.g. 7e+06 vs 7000000, 0.24844704792862 vs 0.24844704792861999).
    Non-numeric input is returned unchanged.
    """
    try:
        return f"{round(float(s), 14):.14g}"
    except ValueError:
        return s


def _fix_gridspan(root: etree._Element) -> None:
    """Step 10: Clamp <a:tc gridSpan> values that exceed the table's column count.

    Shared by structural_normalize (Step 10) and to_normalized_lines_lite.
    """
    for tbl in root.findall(f".//{_A_TBL_TAG}"):
        tbl_grid = tbl.find(_A_TBL_GRID_TAG)
        if tbl_grid is None:
            continue
        col_count = len(tbl_grid.findall(_A_GRID_COL_TAG))
        if col_count == 0:
            continue
        for tc in tbl.findall(f".//{_A_TC_TAG}"):
            gs = tc.get("gridSpan")
            if gs is not None and int(gs) > col_count:
                tc.set("gridSpan", str(col_count))


# ---------------------------------------------------------------------------
# structural_normalize rules — each takes the tree root and mutates it in place.
#
# The "Step N" numbers in the docstrings below are stable identifiers referenced
# from docs/normalization-dev-guide.md. They document *which* rule is which; the
# execution order is defined solely by _NORMALIZE_RULES at the end of this
# section. Adding a rule means appending a new number, never renumbering.
# ---------------------------------------------------------------------------

def _remove_a16_row_col_ids(root: etree._Element) -> None:
    """Step 1: Remove a16:rowId / a16:colId containing a:ext elements."""
    for noisy_tag in (f"{{{_A16_NS}}}rowId", f"{{{_A16_NS}}}colId"):
        for el in root.findall(f".//{noisy_tag}"):
            ext = el.getparent()
            if ext is not None and ext.tag == _A_EXT_TAG:
                extLst = ext.getparent()  # noqa: N806
                if extLst is not None:
                    extLst.remove(ext)


def _remove_empty_ext_and_extlst(root: etree._Element) -> None:
    """Step 2a/2b: Remove empty a:ext then empty a:extLst elements."""
    for ext_el in root.findall(f".//{_A_EXT_TAG}"):
        if len(ext_el) == 0 and not (ext_el.text and ext_el.text.strip()):
            parent = ext_el.getparent()
            if parent is not None and parent.tag == _A_EXT_LST_TAG:
                parent.remove(ext_el)

    for extLst_el in root.findall(f".//{_A_EXT_LST_TAG}"):  # noqa: N806
        if len(extLst_el) == 0 and not (extLst_el.text and extLst_el.text.strip()):
            _detach(extLst_el)


def _remove_empty_text_runs(root: etree._Element) -> None:
    """Step 3: Remove empty text runs <a:r><a:t/></a:r>."""
    for r_el in root.findall(f".//{_A_R_TAG}"):
        t_els = r_el.findall(_A_T_TAG)
        if len(t_els) == 1 and not t_els[0].text:
            _detach(r_el)


def _remove_end_para_rpr(root: etree._Element) -> None:
    """Step 4: Remove <a:endParaRPr> (auto-deleted by PowerPoint on save)."""
    _remove_all(root, _A_END_PARA_RPR_TAG)


def _remove_c_lang(root: etree._Element) -> None:
    """Step 5: Remove <c:lang> (locale-dependent)."""
    _remove_all(root, _C_LANG_TAG)


def _strip_def_rpr_lang(root: etree._Element) -> None:
    """Step 6: Remove lang/altLang from <a:defRPr>."""
    for el in root.findall(f".//{_A_DEF_RPR_TAG}"):
        el.attrib.pop("lang", None)
        el.attrib.pop("altLang", None)


def _remove_chart_solid_fill(root: etree._Element) -> None:
    """Step 7: Remove all <a:solidFill> inside <c:spPr>."""
    for spPr in root.findall(f".//{_C_SPR_TAG}"):  # noqa: N806
        for fill in list(spPr.findall(f".//{_A_SOLID_FILL}")):
            _detach(fill)


def _remove_ser_order(root: etree._Element) -> None:
    """Step 8: Remove <c:order> from <c:ser> (1-based vs 0-based)."""
    for ser in root.findall(f".//{_C_SER_TAG}"):
        for order in list(ser.findall(_C_ORDER_TAG)):
            ser.remove(order)


def _remove_auto_update(root: etree._Element) -> None:
    """Step 9: Remove <c:autoUpdate> from <c:externalData>."""
    for ed in root.findall(f".//{_C_EXT_DATA_TAG}"):
        for au in list(ed.findall(_C_AUTO_UPD_TAG)):
            ed.remove(au)


def _merge_paragraph_runs(root: etree._Element) -> None:
    """Step 11: Merge multiple <a:r> runs in a paragraph into a single run.

    PowerPoint may split/merge runs on save, inserting spaces between them.
    """
    for p_el in root.findall(f".//{_A_P_TAG}"):
        r_els = [child for child in p_el if child.tag == _A_R_TAG]
        if not r_els:
            continue
        if len(r_els) == 1:
            t_el = r_els[0].find(_A_T_TAG)
            if t_el is not None and t_el.text:
                t_el.text = t_el.text.replace(" ", "")
            continue
        # Multiple runs: combine text, discard run properties, keep as single run
        combined = "".join((r.findtext(_A_T_TAG) or "").replace(" ", "") for r in r_els)
        first_r = r_els[0]
        for r in r_els[1:]:
            p_el.remove(r)
        for child in list(first_r):
            first_r.remove(child)
        if combined:
            new_t = etree.SubElement(first_r, _A_T_TAG)
            new_t.text = combined
        else:
            p_el.remove(first_r)


def _clear_auto_updating_fields(root: etree._Element) -> None:
    """Step 12: Clear <a:t> content inside slide-number and date/time <a:fld> fields.

    PowerPoint stores the computed slide number and the date/time of the save as
    the field text, so it changes between file versions.
    """
    for fld in root.findall(f".//{_A_FLD_TAG}"):
        fld_type = fld.get("type", "")
        if fld_type == "slidenum" or fld_type.startswith("datetime"):
            for t_el in fld.findall(f".//{_A_T_TAG}"):
                t_el.text = ""


def _strip_rpr_lang(root: etree._Element) -> None:
    """Step 13: Remove lang/altLang from <a:rPr>, then drop empty <a:rPr/>."""
    for el in root.findall(f".//{_A_RPR_TAG}"):
        el.attrib.pop("lang", None)
        el.attrib.pop("altLang", None)
    for el in list(root.findall(f".//{_A_RPR_TAG}")):
        if len(el) == 0 and not el.attrib:
            _detach(el)


def _remove_invisible_table_borders(root: etree._Element) -> None:
    """Step 14: Remove <a:lnT/B/L/R> borders with bg1 or FFFFFF (invisible) fills."""
    for tcPr in root.findall(f".//{_A_TC_PR_TAG}"):  # noqa: N806
        for border_tag in _A_TABLE_BORDER_TAGS:
            for border in list(tcPr.findall(border_tag)):
                is_invisible = (
                    any(sc.get("val") == "bg1" for sc in border.findall(f".//{_A_SCHEME_CLR_TAG}"))
                    or any(sc.get("val", "").upper() == "FFFFFF" for sc in border.findall(f".//{_A_SRGB_CLR_TAG}"))
                )
                if is_invisible:
                    tcPr.remove(border)


def _strip_tc_hmerge(root: etree._Element) -> None:
    """Step 15: Remove hMerge attribute from <a:tc>."""
    for tc in root.findall(f".//{_A_TC_TAG}"):
        tc.attrib.pop("hMerge", None)


def _strip_external_data_rid(root: etree._Element) -> None:
    """Step 16: Remove r:id from <c:externalData> (rIds get renumbered)."""
    for ed in root.findall(f".//{_C_EXT_DATA_TAG}"):
        ed.attrib.pop(f"{{{_R_NS}}}id", None)


def _normalize_axis_floats(root: etree._Element) -> None:
    """Step 17: Normalize float notation in chart axis val attributes."""
    for tag in _C_AXIS_FLOAT_TAGS:
        for el in root.findall(f".//{tag}"):
            val = el.get("val")
            if val:
                el.set("val", _normalize_float_str(val))


def _remove_null_data_points(root: etree._Element) -> None:
    """Step 18: Remove <c:pt> elements containing only an empty <c:v/>."""
    for pt in list(root.findall(f".//{_C_PT_TAG}")):
        children = list(pt)
        if len(children) == 1 and children[0].tag == _C_V_TAG and not children[0].text:
            _detach(pt)


def _normalize_value_floats(root: etree._Element) -> None:
    """Step 19: Normalize <c:v> floating-point notation (chart data values)."""
    for el in root.findall(f".//{_C_V_TAG}"):
        if el.text:
            el.text = _normalize_float_str(el.text)


def _remove_auto_data_point_overrides(root: etree._Element) -> None:
    """Step 20: Remove PowerPoint auto-generated <c:dPt> overrides.

    Two cases: (a) contains <c16:uniqueId> in extLst, (b) only <c:idx> + empty <c:spPr>.
    """
    for dpt in list(root.findall(f".//{_C_DPT_TAG}")):
        has_unique_id = any(True for _ in dpt.findall(f".//{_C16_UNIQUE_ID_TAG}"))
        children = list(dpt)
        is_empty_override = (
            not has_unique_id
            and len(children) == 2
            and children[0].tag == _C_IDX_TAG
            and children[1].tag == _C_SPR_TAG
            and len(children[1]) == 0
        )
        if has_unique_id or is_empty_override:
            _detach(dpt)


def _sort_scaling_children(root: etree._Element) -> None:
    """Step 21: Sort <c:scaling> children by tag name."""
    for scaling in root.findall(f".//{_C_SCALING_TAG}"):
        scaling[:] = sorted(scaling, key=lambda c: c.tag)


def _strip_pt_format_code(root: etree._Element) -> None:
    """Step 22: Remove formatCode attribute from <c:pt> elements."""
    for pt in root.findall(f".//{_C_PT_TAG}"):
        pt.attrib.pop("formatCode", None)


def _normalize_num_cache_format(root: etree._Element) -> None:
    """Step 23: Normalize <c:formatCode> inside <c:numCache> to "General"."""
    for nc in root.findall(f".//{_C_NUM_CACHE_TAG}"):
        for fc in nc.findall(_C_FORMAT_CODE_TAG):
            fc.text = "General"


def _remove_num_fmt(root: etree._Element) -> None:
    """Step 24: Remove <c:numFmt> from chart elements."""
    _remove_all(root, _C_NUM_FMT_TAG)


def _remove_chart_txpr(root: etree._Element) -> None:
    """Step 25: Remove <c:txPr> (auto-generated text body props) from charts."""
    _remove_all(root, _C_TX_PR_TAG)


def _remove_axis_label_flags(root: etree._Element) -> None:
    """Step 26: Remove <c:lblAlgn> and <c:noMultiLvlLbl> from chart axes."""
    _remove_all(root, _C_LBL_ALGN_TAG)
    _remove_all(root, _C_NO_MULTI_LVL_TAG)


def _remove_auto_zero_crosses(root: etree._Element) -> None:
    """Step 27: Remove <c:crosses val="autoZero"> from chart axes."""
    for el in list(root.findall(f".//{_C_CROSSES_TAG}")):
        if el.get("val") == "autoZero":
            _detach(el)


# Applied top to bottom. This ordering is the contract; the Step numbers are
# only labels that let the dev guide refer to an individual rule.
_NORMALIZE_RULES = (
    _remove_a16_row_col_ids,          # Step 1
    _remove_empty_ext_and_extlst,     # Step 2a/2b
    _remove_empty_text_runs,          # Step 3
    _remove_end_para_rpr,             # Step 4
    _remove_c_lang,                   # Step 5
    _strip_def_rpr_lang,              # Step 6
    _remove_chart_solid_fill,         # Step 7
    _remove_ser_order,                # Step 8
    _remove_auto_update,              # Step 9
    _fix_gridspan,                    # Step 10
    _merge_paragraph_runs,            # Step 11
    _clear_auto_updating_fields,      # Step 12
    _strip_rpr_lang,                  # Step 13
    _remove_invisible_table_borders,  # Step 14
    _strip_tc_hmerge,                 # Step 15
    _strip_external_data_rid,         # Step 16
    _normalize_axis_floats,           # Step 17
    _remove_null_data_points,         # Step 18
    _normalize_value_floats,          # Step 19
    _remove_auto_data_point_overrides,  # Step 20
    _sort_scaling_children,           # Step 21
    _strip_pt_format_code,            # Step 22
    _normalize_num_cache_format,      # Step 23
    _remove_num_fmt,                  # Step 24
    _remove_chart_txpr,               # Step 25
    _remove_axis_label_flags,         # Step 26
    _remove_auto_zero_crosses,        # Step 27
)


def structural_normalize(xml_str: str) -> str:
    """
    Strong normalization applied before Jaccard mapping.

    Removes PowerPoint-generated structural noise so that semantically identical
    slides have identical (or near-identical) XML tokens. Each transformation
    lives in its own rule function; see _NORMALIZE_RULES for the ordered list.

    Uses remove_blank_text=True to strip whitespace-only text nodes, ensuring
    consistent output regardless of input indentation style (minidom vs lxml).

    Used by: slide-compare-lite (and by slide-compare for mapping only)
    """
    try:
        root = etree.fromstring(xml_str.encode("utf-8"), parser=_BLANK_TEXT_PARSER)
    except etree.XMLSyntaxError:
        return xml_str

    for rule in _NORMALIZE_RULES:
        rule(root)

    # Final: Sort all attributes alphabetically.
    # Attribute order in raw XML can differ between files even when values match.
    _sort_attrs_recursive(root)

    return etree.tostring(root, encoding="unicode")


def _sort_attrs_recursive(el: etree._Element) -> None:
    """Sort attributes alphabetically on every element in the subtree."""
    sorted_attrs = sorted(el.attrib.items())
    el.attrib.clear()
    for key, val in sorted_attrs:
        el.attrib[key] = val
    for child in el:
        _sort_attrs_recursive(child)


def normalize_element_lite(el: etree._Element) -> None:
    """
    Normalize a single XML element in-place for lite diff generation.

    Applies all PowerPoint-aware normalizations on top of structural normalization.
    Used by: slide-compare-lite differ
    """
    if el.text:
        el.text = el.text.strip() or None
    if el.tail:
        el.tail = el.tail.strip() or None

    # Normalize c:v floating-point notation
    if el.tag == _C_V_TAG and el.text:
        el.text = _normalize_float_str(el.text)

    # Sort c:scaling children by tag name
    if el.tag == _C_SCALING_TAG:
        el[:] = sorted(el, key=lambda c: c.tag)

    # Sort chart type container children (keep c:ser first)
    if el.tag in _CHART_TYPE_CONTAINERS:
        ser_children   = [c for c in el if c.tag == _C_SER_TAG]
        other_children = sorted([c for c in el if c.tag != _C_SER_TAG], key=lambda c: c.tag)
        el[:] = ser_children + other_children

    # Remove <c:order> from <c:ser>
    if el.tag == _C_SER_TAG:
        for order in list(el.findall(_C_ORDER_TAG)):
            el.remove(order)

    # Remove <c:autoUpdate> from <c:externalData>
    if el.tag == _C_EXT_DATA_TAG:
        for au in list(el.findall(_C_AUTO_UPD_TAG)):
            el.remove(au)

    # Sort <c:chart> children by tag name
    if el.tag == _C_CHART_TAG:
        el[:] = sorted(el, key=lambda c: c.tag)

    # Remove lang/altLang from <a:defRPr>
    if el.tag == _A_DEF_RPR_TAG:
        el.attrib.pop("lang", None)
        el.attrib.pop("altLang", None)

    # Sort attributes alphabetically
    sorted_attrs = sorted(el.attrib.items())
    el.attrib.clear()
    for key, val in sorted_attrs:
        el.attrib[key] = val

    for child in el:
        normalize_element_lite(child)


def normalize_element_minimal(el: etree._Element) -> None:
    """
    Minimal normalization for full diff generation.

    Only normalizes whitespace, attribute order, and floating-point notation.
    Used by: slide-compare differ
    """
    if el.text:
        el.text = el.text.strip() or None
    if el.tail:
        el.tail = el.tail.strip() or None

    # Normalize c:v floating-point notation
    if el.tag == _C_V_TAG and el.text:
        el.text = _normalize_float_str(el.text)

    # Sort attributes alphabetically
    sorted_attrs = sorted(el.attrib.items())
    el.attrib.clear()
    for key, val in sorted_attrs:
        el.attrib[key] = val

    for child in el:
        normalize_element_minimal(child)


def to_normalized_lines_lite(xml_str: str) -> list[str]:
    """Parse XML, apply lite normalization, return pretty-printed lines."""
    root = etree.fromstring(xml_str.encode("utf-8"))
    # Remove p14 noisy tags
    for tag in _P14_NOISY_TAGS:
        _remove_all(root, tag)
    # Remove c:lang
    _remove_all(root, _C_LANG_TAG)
    _fix_gridspan(root)
    normalize_element_lite(root)
    return etree.tostring(root, pretty_print=True).decode().splitlines(keepends=True)


def to_normalized_lines_minimal(xml_str: str) -> list[str]:
    """Parse XML, apply minimal normalization, return pretty-printed lines."""
    root = etree.fromstring(xml_str.encode("utf-8"))
    normalize_element_minimal(root)
    return etree.tostring(root, pretty_print=True).decode().splitlines(keepends=True)


# ---------------------------------------------------------------------------
# Node-serialization for slide-compare (one line per XML node)
# Avoids line-number drift caused by XML tag insertion/deletion.
# ---------------------------------------------------------------------------

_NS_PREFIXES = {
    "http://schemas.openxmlformats.org/presentationml/2006/main": "p",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "a",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "c",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "xdr",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships": "r",
    "http://schemas.microsoft.com/office/powerpoint/2010/main": "p14",
    "http://schemas.microsoft.com/office/drawing/2014/main": "a16",
}


def _short_tag(tag: str) -> str:
    if tag.startswith("{"):
        ns, local = tag[1:].split("}", 1)
        prefix = _NS_PREFIXES.get(ns, ns.split("/")[-1])
        return f"{prefix}:{local}"
    return tag


def xml_to_node_lines(root: etree._Element) -> list[str]:
    """
    Serialize XML tree to one-line-per-node format.

    Each element is represented as a single line:
      DEPTH  <tag attr="val">text
    This ensures that adding/removing a node affects exactly one line
    in the diff, regardless of the element's position in the tree.
    """
    lines = []

    def walk(el: etree._Element, depth: int = 0) -> None:
        tag = _short_tag(el.tag)
        attrs = "".join(f' {_short_tag(k)}="{v}"' for k, v in sorted(el.attrib.items()))
        text = (el.text or "").strip()
        line = "  " * depth + f"<{tag}{attrs}>"
        if text:
            line += text
        lines.append(line + "\n")
        for child in el:
            walk(child, depth + 1)

    walk(root)
    return lines


def to_node_lines_minimal(xml_str: str) -> list[str]:
    """Parse XML, apply minimal normalization, return one-line-per-node."""
    root = etree.fromstring(xml_str.encode("utf-8"))
    normalize_element_minimal(root)
    return xml_to_node_lines(root)
