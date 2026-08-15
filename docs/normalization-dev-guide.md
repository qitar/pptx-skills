# Normalization Development Guide

How to add a new rule to `structural_normalize()`.

This is the procedure for the situation where slides are reported as `modified` even though
nothing was edited — meaning a PowerPoint save artifact is not being normalized away yet.

> **Prerequisite:** you need PowerPoint itself for this work. What these rules remove is XML that
> PowerPoint rewrites when it saves, and PowerPoint is the only thing that reproduces those
> artifacts. Running the comparison skill does not require PowerPoint.

---

## 0. Prepare a pair of test decks

This repository ships no PPTX files — `.gitignore` excludes `*.pptx` so that real decks are never
committed. Build the two kinds of pair below yourself. Any working directory will do; `fixtures/`
is already gitignored.

### Pair A — the unchanged pair (finds false positives)

This is your measure of whether normalization is working. **Every slide should come out
`identical`.**

1. Open any deck in PowerPoint and save it as `fixtures/before.pptx`
2. Close PowerPoint
3. Reopen `fixtures/before.pptx` and, **without editing anything**, save it as `fixtures/after.pptx`

The result is a pair whose content is identical but whose XML PowerPoint has rewritten. Whatever
comes out `modified` is exactly your list of missing normalization.

### Pair B — the known-change pair (finds over-removal)

This is your measure of whether normalization has become so aggressive that it erases real edits.
**Only the changes you made intentionally should come out `modified`.**

1. Copy `fixtures/before.pptx` to `fixtures/changed.pptx`
2. Write down what you are about to change (e.g. the title on slide 3, one chart value on slide 7)
3. Make only those changes and save

One pair that touches a textbox, a table, a chart and a speaker note is enough to cover the
common cases in a single run.

> To show that a new rule removes only artifacts and keeps real edits, you must check **both
> pairs**. Pair A alone cannot detect over-removal.

---

## 1. Detect the problem

Compare pair A with `lite` and count the false positives.

```bash
skills/pptx-compare/run.sh lite fixtures/before.pptx fixtures/after.pptx
```

**What to look for:**

- `summary.identical` lower than the total slide count — the difference is un-normalized noise
- On pair B, only the changes recorded in step 0 should be `modified`

---

## 2. Locate what still differs

A throwaway diagnostic script that shows which slide differs and where. Run it from the
repository root.

```python
# debug_remaining.py (a scratch file; no need to commit it)
import sys, difflib
from pathlib import Path

SKILL_DIR = Path("skills/pptx-compare")
sys.path.insert(0, str(SKILL_DIR))

from core.extractor import extract_slides
from core.mapper import match_slides

BEFORE = "fixtures/before.pptx"
AFTER  = "fixtures/after.pptx"

before_slides, _ = extract_slides(BEFORE, apply_normalization=True)
after_slides,  _ = extract_slides(AFTER,  apply_normalization=True)

matches = match_slides(before_slides, after_slides, threshold=0.5)
before_by_idx = {s["index"]: s for s in before_slides}
after_by_idx  = {s["index"]: s for s in after_slides}

# Treat byte-identical XML as identical
for m in matches:
    if m.status == "modified":
        if before_by_idx[m.before_index]["xml_content"] == after_by_idx[m.after_index]["xml_content"]:
            m.status = "identical"

# Print the first differing position for each modified slide
for m in sorted(matches, key=lambda x: x.before_index or 0):
    if m.status != "modified":
        continue
    b_xml = before_by_idx[m.before_index]["xml_content"]
    a_xml = after_by_idx[m.after_index]["xml_content"]
    b_lines, a_lines = b_xml.splitlines(), a_xml.splitlines()
    diffs = list(difflib.unified_diff(b_lines, a_lines, lineterm="", n=0))

    for bl, al in zip(b_lines, a_lines):
        if bl != al:
            j = next((j for j in range(min(len(bl), len(al))) if bl[j] != al[j]), min(len(bl), len(al)))
            print(f"\nB{m.before_index} (sim={m.similarity:.4f}, {len(diffs)} diff lines):")
            print(f"  B: {repr(bl[max(0,j-40):j+100])}")
            print(f"  A: {repr(al[max(0,j-40):j+100])}")
            break
```

It needs lxml, so use the venv the skill already created:

```bash
skills/pptx-compare/.venv/bin/python debug_remaining.py            # macOS / Linux
skills\pptx-compare\.venv\Scripts\python.exe debug_remaining.py    # Windows
```

### Reading the output

- `sim=1.0000` but still `modified` — the Jaccard tokens match while the strings do not, so this
  is an attribute-ordering or whitespace problem
- Many diff lines (50+) — a repeating pattern, typically table borders or chart data
- Few diff lines (1-5) — one specific element differs in one place, usually easy to fix

---

## 3. Decide what kind of difference it is

### Artifact, or a real change?

| Signal | Artifact | Real change |
|---|---|---|
| **Visual effect** | None (background colour = bg1, invisible borders) | Visible (text content, data values, structure) |
| **Reproducibility** | Always appears in the same shape | Content differs slide by slide |
| **Origin** | A formatting difference PowerPoint writes on save | An edit the user intended |
| **Spread across similar slides** | The same pattern on many slides | Only on specific slides |

When in doubt, check whether it shows up in pair A. **Anything that appears in the unchanged pair
is an artifact by definition.**

### Known PowerPoint artifact patterns

The rules below are already implemented in `structural_normalize()`. Use them as a reference when
you hit something similar.

| XML pattern | What it is | How it is handled |
|---|---|---|
| `<a:lnT/B/L/R>` with `schemeClr val="bg1"` | Invisible table border | Remove `lnT/B/L/R` from `<a:tcPr>` |
| `<a:lnT/B/L/R>` with `srgbClr val="FFFFFF"` | Invisible white border | Same as above |
| `<c:dPt>` containing `<c16:uniqueId>` | Data-point formatting PowerPoint adds automatically | Remove the whole `<c:dPt>` |
| `<c:dPt>` of just `<c:idx>` + empty `<c:spPr>` | Same, without the c16 extension | Remove the whole `<c:dPt>` |
| Float notation in `<c:v>` | `1.531e+06` vs `1531000`, trailing-digit drift | Canonicalize to `float(v):.14g` |
| `val` notation on `<c:min>/<c:max>` | Same, for axis scales | Canonicalize to `float(val):.14g` |
| Child order inside `<c:scaling>` | `<c:min>` and `<c:max>` swap places | Sort children by tag name |
| Added `<c:numFmt>` | PowerPoint adds a format to the axis | Remove the whole element |
| `<c:formatCode>` in `<c:numCache>` | `0.0%` vs `General` | Normalize to `General` |
| `<c:pt formatCode="...">` | PowerPoint stamps a format on each point | Drop the `formatCode` attribute |
| `<c:txPr>` in chart axes or `<c:dLbls>` | Text body properties PowerPoint adds | Remove the whole element |
| `<c:lblAlgn>`, `<c:noMultiLvlLbl>` | Axis-label settings PowerPoint adds | Remove the whole element |
| `<a:rPr lang="ja-JP" altLang="en-US">` | Language attributes PowerPoint adds on save | Drop `lang` / `altLang` |
| `<a:fld type="slidenum"><a:t>N</a:t>` | PowerPoint precomputes and stores the page number | Clear the `<a:t>` text |
| `<a:tc hMerge="1">` | Merged-cell attribute PowerPoint drops | Drop the `hMerge` attribute |
| `<c:autoUpdate>` in `<c:externalData>` | External-data auto-refresh flag | Remove the whole element |
| `<c:lang>` | Locale-dependent language tag | Remove the whole element |
| `<a:endParaRPr>` | PowerPoint deletes this on save | Remove the whole element |

---

## 4. Add the rule to `normalizer.py`

File: `skills/pptx-compare/core/normalizer.py`

Each normalization rule is a standalone function, and they are applied in the order they appear in
the `_NORMALIZE_RULES` tuple. The `Step N` numbers in the docstrings are stable identifiers so
this guide can refer to an individual rule; the execution order comes from the tuple, not from the
numbers. **Never renumber existing rules — append the next number.**

### Steps

**a. Add any tag constants you need** (the constants section near the top of the file)

```python
_C_NEW_TAG = f"{{{_C_NS}}}newElement"
```

**b. Add the rule function**

```python
def _remove_new_element(root: etree._Element) -> None:
    """Step N: <what it removes>.

    <one line on why this is an artifact>
    """
    _remove_all(root, _C_NEW_TAG)
```

**c. Register it in `_NORMALIZE_RULES`**

Append it at the end (after `_remove_auto_zero_crosses`) as a rule. Insert it earlier only if it
depends on the output of a preceding rule.

```python
_NORMALIZE_RULES = (
    ...
    _remove_auto_zero_crosses,        # Step 27
    _remove_new_element,              # Step 28
)
```

The final attribute sort (`_sort_attrs_recursive`) runs in `structural_normalize()` after every
rule, so individual rules do not need to think about it.

### Common patterns

```python
# Remove an element entirely (helper available)
_remove_all(root, _TAG)

# Drop an attribute
for el in root.findall(f".//{_TAG}"):
    el.attrib.pop("attrName", None)

# Canonicalize an attribute value
for el in root.findall(f".//{_TAG}"):
    val = el.get("val")
    if val:
        el.set("val", _normalize_float_str(val))

# Clear text content
for el in root.findall(f".//{_TAG}"):
    if el.get("type") == "someType":
        for child in el.findall(f".//{_A_T_TAG}"):
            child.text = ""

# Sort children by tag name
for el in root.findall(f".//{_CONTAINER_TAG}"):
    el[:] = sorted(el, key=lambda c: c.tag)

# Conditional removal, based on what the children look like
for el in list(root.findall(f".//{_PARENT_TAG}")):
    children = list(el)
    if len(children) == 2 and children[0].tag == _IDX_TAG and len(children[1]) == 0:
        _detach(el)
```

### Things to watch for

- **Check the parent element.** `<a:lnT>` is a child of `<a:tcPr>`, not a direct child of `<a:tc>`.
- **Copy the list before removing.** `for el in list(root.findall(...)):` is safe to delete from
  while iterating.
- **Keep the condition narrow.** Only remove what you are confident is an artifact. `slide-compare`
  applies minimal normalization, so anything you are unsure about can still be inspected there.

---

## 5. Verify

Re-run the command from step 1 against **both pair A and pair B**.

```bash
skills/pptx-compare/run.sh lite fixtures/before.pptx  fixtures/after.pptx    # pair A
skills/pptx-compare/run.sh lite fixtures/before.pptx  fixtures/changed.pptx  # pair B
```

**Checklist:**

| Target | Expected result |
|---|---|
| Pair A | `identical` has gone up (if it went down, the rule is broken) |
| Pair B | The changes recorded in step 0 are still reported as `modified` |
| Any other pairs you keep | `identical` has not gone down |

If a change in pair B disappeared, the rule is **over-removing**. Narrow its condition.

Normalization should prioritise never erasing a real change over driving pair A to 100%. Missing a
genuine edit is far more harmful than leaving a little noise behind. When an artifact is
borderline, the safe call is to leave it out of `lite` and let people inspect it through
`slide-compare`, which applies minimal normalization.

### After adding a rule

Add a row to the "Known PowerPoint artifact patterns" table above. It is what the next person hits
the same symptom will search, and it saves them re-deriving a rule that already exists.
