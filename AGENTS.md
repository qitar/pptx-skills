# AGENTS.md

This file provides guidance to AI coding agents and human contributors working on the
`slidepack-skills` codebase.

## Project Overview

`slidepack-skills` packages a Claude Skill that diffs two PowerPoint files at the OOXML level.
The hard part is not diffing XML — it is discarding the XML that PowerPoint rewrites on every
save. Most of the domain knowledge in this repo lives in `core/normalizer.py`.

## Commands

| Command | Description |
| --- | --- |
| `skills/pptx-compare/run.sh lite <before> <after>` | Fast pass: what changed |
| `skills/pptx-compare/run.sh full <before> <after>` | Detailed pass: how it changed |
| `skills\pptx-compare\run.ps1` | The same two commands, for Windows PowerShell |
| `ruff check .` | Lint |

The launchers create `skills/pptx-compare/.venv` on first run and install `requirements.txt`
into it. There is no repo-level install step, no build step, no test suite and no CI.

## Architecture

```
skills/pptx-compare/            the distributable skill
├── SKILL.md                    entry point — picks a strategy
├── run.sh / run.ps1            venv bootstrap + dispatch
├── requirements.txt            runtime deps, installed into the skill's own venv
├── core/
│   ├── extractor.py            unzip pptx, inline charts + notes, read CRC32
│   ├── normalizer.py           the normalization rules (Step 1..N)
│   ├── mapper.py               Jaccard + position-biased slide matching
│   ├── differ.py               unified / node-based diff writing
│   └── utils.py                pretty-print, rels resolution, diff output
├── slide-compare/              strategy: minimal normalization, node diff
│   ├── slide-compare.md        workflow + report format for Claude
│   └── scripts/main.py         CLI entry point
└── slide-compare-lite/         strategy: strong normalization, unified diff
    ├── slide-compare-lite.md   workflow + report format for Claude
    └── scripts/main.py         CLI entry point

docs/normalization-dev-guide.md  how to diagnose an artifact and add a rule
```

`skills/pptx-compare/` is the unit of distribution. Installing means copying (or symlinking) that
directory into `~/.claude/skills/` or a project's `.claude/skills/` — there is no packaging step,
and nothing outside that directory ships with the skill. Anything a user needs at runtime must
live inside it.

### Data flow

```
pptx (zip)
  → extractor    unzip, pretty-print, inline chart XML and notes XML into each slide
  → normalizer   structural_normalize(): strip PowerPoint save-noise
  → mapper       score every before/after pair, greedy 1:1 assignment
  → differ       write one diff file per modified pair
  → main.py      emit JSON: summary + mapping + diff paths
```

`lite` runs this once with strong normalization and a CRC32 pre-pass.
`full` runs extraction four times: twice with minimal normalization (for the diff) and twice
with strong normalization (for the mapping), because Jaccard matching is unreliable on
un-normalized XML.

## Conventions

### Normalization rules

- Each rule is a standalone function taking the tree root and mutating it in place.
- `_NORMALIZE_RULES` defines the execution order. That tuple is the contract.
- The `Step N` numbers in docstrings are stable labels so the dev guide can refer to a rule.
  **Never renumber existing rules** — append the next number.
- Prefer narrow conditions. A rule that removes a real user edit is much worse than a rule that
  leaves some noise behind; `slide-compare` exists to show what `lite` normalized away.
- Read `docs/normalization-dev-guide.md` before adding one.

### Dependencies

`skills/pptx-compare/requirements.txt` is what ships with the skill; `[project].dependencies` in
`pyproject.toml` is what contributors read. Nothing checks that they agree, so update both in the
same commit. Keep the dependency set minimal — `lxml` is the only runtime dependency, and the
skill should stay installable without a compiler toolchain where possible.

### Line endings

`.gitattributes` pins `*.sh` to LF in the working tree, not just in the repository. Because the
skill directory is copied verbatim into `~/.claude/skills/`, `run.sh` is executed exactly as it
appears in the repo — a CRLF shebang committed from Windows breaks every Linux and macOS user
with `bad interpreter: /usr/bin/env bash^M`. If you edit a shell script with a tool that rewrites
line endings, check `git ls-files --eol` before committing.

### PowerShell

`run.ps1` deliberately sets `$ErrorActionPreference = 'Continue'`, not `'Stop'`. In Windows
PowerShell 5.1 any stderr output from a native executable is wrapped in a `NativeCommandError`,
which under `'Stop'` aborts the script even on success — pip warnings and the Microsoft Store
python stub both trigger this. Exit codes are checked explicitly instead.

### Licensing

`LICENSE` is the Apache-2.0 text verbatim, including the untouched `[yyyy] [name of copyright
owner]` template in its appendix — do not fill that in. The real attribution lives in `NOTICE`,
which is what section 4(d) requires redistributions to carry. Update `NOTICE`, not `LICENSE`.

Because a skill is installed by copying `skills/pptx-compare/` on its own, neither file travels
with it; the `license:` field in `SKILL.md` frontmatter is what identifies the license at the
destination. `license` is one of the fields the skill spec allows, alongside `name`,
`description`, `compatibility`, `metadata` and `allowed-tools` — adding any field outside that
set makes claude.ai packaging and upload fail outright.

### Language

Prompts are Japanese; everything else is English. `SKILL.md` and the two strategy documents are
written in Japanese on purpose — they are prompts, and they make Claude report in Japanese by
default. Everything a person reads (README, this file, `docs/`) and everything a machine reads
(code, comments, JSON keys, diff output) is English.

Do not translate a prompt file into English as a "consistency" cleanup; that changes the language
Claude reports in. When editing a strategy document, keep the report-format templates intact —
Claude follows them literally.

## Testing

There is no automated test suite and no CI. Verification is manual and requires PowerPoint,
because the artifacts being normalized are produced by PowerPoint's save path.
`docs/normalization-dev-guide.md` describes how to build the before/after pairs needed to
validate a change.

This is a deliberate choice rather than an omission. Most of what this repository ships is prompt
text that no test can meaningfully assert on, and the Python underneath it cannot be exercised
without real PowerPoint output. Before submitting a change, run `ruff check .` and compare a
before/after pair by hand.
