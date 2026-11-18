# pptx-skills/pptx-compare

A [Claude Skill](https://code.claude.com/docs/en/skills) for diffing PowerPoint (`.pptx`) files
at the OOXML level.

Comparing two pptx decks is not straightforward: PowerPoint rewrites parts of the XML every time a
file is saved, regardless of intentional user edits. A naive XML diff drowns
in that noise, while a text-extraction diff misses formatting changes entirely.

This skill normalizes away the noise, maps slide pairs by similarity
(so reordering, insertion and deletion are handled), and reports what changed.

## Install

Recommended:

```
npx skills add github.com/qitar/pptx-skills
```

Manual: Clone this repo, then

```
ln -s skills/pptx-compare ~/.agents/skills
```

## Requirements

**Python >= 3.10** in your path.

Script dependencies get installed to a venv within the skill directory on first run.

PowerPoint itself is not required.

## Skill usage

Ask Claude to compare two decks:

> Compare before.pptx and after.pptx

Two strategies are available:

| | `slide-compare-lite` | `slide-compare` |
| --- | --- | --- |
| Purpose | **Which** slides changed | **What** changed on each slide |
| Normalization | Strong (removes PowerPoint save noise) | Minimal (attribute sort / whitespace only) |
| Speed | Fast (less tokens) | Slow (more tokens) |

Start with `lite` to find out *whether* something changed, then run the full pass if they did.

Slide masters, layouts and themes are compared as well as the slides themselves.


## Standalone script usage

The skill comes with a plain Python program that you can run without Claude.

```bash
# macOS / Linux / Git Bash
skills/pptx-compare/run.sh lite before.pptx after.pptx
skills/pptx-compare/run.sh full before.pptx after.pptx
```

```powershell
# Windows PowerShell — same arguments, no bash required
skills\pptx-compare\run.ps1 lite before.pptx after.pptx
skills\pptx-compare\run.ps1 full before.pptx after.pptx
```

Both print JSON to stdout — a per-slide mapping with `identical` / `modified` / `added` / `deleted`
status, plus paths to the generated diff files.

| Option | Default | Description |
| --- | --- | --- |
| `--threshold` | `0.5` | Minimum Jaccard similarity to treat two slides as the same when mapping pairs |
| `--output-dir DIR` | temp dir | Where diff files are written |

## Repository layout

```
skills/pptx-compare/             the skill itself
├── SKILL.md                  point: strategy selection
├── run.sh / run.ps1          venv bootstrap + dispatch
├── core/                     extractor / normalizer / mapper / differ
├── slide-compare/            full-detail strategy
└── slide-compare-lite/       fast overview strategy

docs/normalization-dev-guide.md   how to add a normalization rule
AGENTS.md                         orientation for AI coding agents
```

The normalizer is where most of the domain knowledge lives. If you find a PowerPoint artifact that
still shows up as a false positive, `docs/normalization-dev-guide.md` walks through diagnosing it
and adding a rule.

## Note on language

Prompts were written in Japanese because our dev team consists of Japanese people. This shouldn't affect usage in other languages.

## Contributing

See [DEVELOPMENT.md](DEVELOPMENT.md) for the development setup, the lint command, and a map of the codebase.

Bug reports and feature requests are welcome via GitHub issues.

## License

Apache License 2.0 ([LICENSE](LICENSE)).

[NOTICE](NOTICE) holds the copyright attribution. Please include that file when redistributing.
