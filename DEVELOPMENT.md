# Development

## Project Overview

`slidepack-skills` packages a Claude Skill that diffs two PowerPoint files at the OOXML level.
Diffing XML is the easy part. The hard part is discarding the XML diff noise,
and most of the domain knowledge for that lives in `core/normalizer.py`.

## Commands

| Command | Description |
| --- | --- |
| `skills/pptx-compare/run.sh lite <before> <after>` | Fast pass: what changed |
| `skills/pptx-compare/run.sh full <before> <after>` | Detailed pass: how it changed |
| `skills\pptx-compare\run.ps1` | The same two commands, for Windows PowerShell |
| `ruff check .` | Lint |

The launchers create `skills/pptx-compare/.venv` on first run and install `requirements.txt`
into it.

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

`skills/pptx-compare/` is the unit of distribution. Installing the skill means copying or
symlinking that directory into `~/.claude/skills/` or a project's `.claude/skills/`. Nothing
outside the directory ships, so anything a user needs at runtime must live inside it.

### Data flow

```
pptx (zip)
  → extractor    unzip, pretty-print, inline chart XML and notes XML into each slide
  → normalizer   structural_normalize(): strip PowerPoint save-noise
  → mapper       score every before/after pair, greedy 1:1 assignment
  → differ       write one diff file per modified pair
  → main.py      emit JSON: summary + mapping + diff paths
```

`lite` runs this pipeline once, with strong normalization and a CRC32 pre-pass.
`full` extracts each file twice: with minimal normalization for the diff, and with strong
normalization for the mapping, because Jaccard matching is unreliable on un-normalized XML.

## Conventions

### Normalization rules

Read `docs/normalization-dev-guide.md` before adding a rule.

- Each rule is a standalone function that takes the tree root and mutates it in place.
- Rules run in the order listed in `_NORMALIZE_RULES`.
- The `Step N` numbers in docstrings are stable labels that the dev guide uses to refer to
  rules. **Never renumber an existing rule.** Give a new rule the next number.
- Prefer narrow conditions. A rule that removes a real user edit is much worse than one that
  leaves some noise behind.

### Dependencies

`lxml` is the only runtime dependency. Keep the set minimal, so that the skill stays installable
without a compiler toolchain.

The runtime dependencies are declared in two places: `skills/pptx-compare/requirements.txt` ships
with the skill, and `[project].dependencies` in `pyproject.toml` is what contributors read.
Keep them in sync.

### Line endings

Shell scripts must have LF line endings in the working tree, not just in the repository. The
skill directory is copied as it is into `~/.claude/skills/`, and a CRLF shebang fails on Linux
and macOS with `bad interpreter: /usr/bin/env bash^M`.

`.gitattributes` pins `*.sh` to LF. If you edit a shell script with a tool that rewrites line
endings, check `git ls-files --eol` before committing.

### PowerShell

Do not change `$ErrorActionPreference` in `run.ps1` from `'Continue'` to `'Stop'`. Windows
PowerShell 5.1 wraps any stderr output from a native executable in a `NativeCommandError`, which
under `'Stop'` aborts the script even when the command succeeded. pip warnings and the Microsoft
Store python stub both trigger this. The script checks exit codes explicitly instead.


### Language

Prompts are in Japanese; everything else is in English.

The prompts are `SKILL.md` and the two strategy documents. Because they are in Japanese, Claude
reports in Japanese by default, so do not translate them into English for consistency. When
editing a strategy document, keep the report-format templates intact, because Claude follows them
literally.

## Testing

Before submitting a change, run `ruff check .` and compare a before/after pair by hand.
`docs/normalization-dev-guide.md` describes how to build the pairs, which requires PowerPoint.

There is no automated test suite and no CI, and this is deliberate. Most of what this repository
ships is prompt text that no test can meaningfully assert on, and the Python underneath it cannot
be exercised without real PowerPoint output.
