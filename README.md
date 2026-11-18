# slidepack-skills

A [Claude Skill](https://code.claude.com/docs/en/skills) for diffing PowerPoint (`.pptx`) files
at the OOXML level.

Comparing two decks is harder than it looks: PowerPoint rewrites parts of the XML every time a
file is opened and saved — language attributes, chart formatting hints, invisible table borders,
float notation, text-run splits — none of which the user actually changed. A naive XML diff drowns
in that noise, while a text-extraction diff misses formatting changes entirely.

This skill normalizes the PowerPoint-generated noise away, maps slides between the two files by
Jaccard similarity (so reordering, insertion and deletion are handled), and reports what changed
in textboxes, tables, charts and speaker notes.

## Two strategies

| | `slide-compare-lite` | `slide-compare` |
| --- | --- | --- |
| Purpose | *What* changed | *How* it changed, at the XML level |
| Normalization | Strong (removes PowerPoint save-noise) | Minimal (attribute sort / whitespace only) |
| Output | File metadata, summary, unified diff | Node-based diff, or raw before/after for small XML |
| Speed | Faster — CRC32 pre-matching skips unchanged parts | Slower — every part is read four times and diffed |

Start with `lite` to find out *whether* something changed, then run the full pass on the decks
that did. Slide masters, layouts and themes are compared as well as the slides themselves.

## Install

A skill is just a directory containing `SKILL.md`, so installing it is a copy. Use it across all
your projects:

```bash
cp -r skills/pptx-compare ~/.claude/skills/
```

…or scope it to one repository:

```bash
cp -r skills/pptx-compare /path/to/your/project/.claude/skills/
```

A symlink works too, if you would rather track this repository:

```bash
ln -s "$(pwd)/skills/pptx-compare" ~/.claude/skills/pptx-compare
```

If you have already run the skill from this checkout, exclude the venv it created so you copy
sources rather than a machine-specific build artifact:

```bash
rsync -a --exclude .venv skills/pptx-compare ~/.claude/skills/
```

Then just ask Claude to compare two decks:

> before.pptx と after.pptx の差分を教えて

> **Uploading to claude.ai instead?** That path — the Skills API, Cowork and cloud sessions —
> takes a zipped skill rather than a directory. Use the official
> [`package_skill.py`](https://github.com/anthropics/skills) to build it; this repository
> deliberately ships no packaging script of its own.

## Standalone usage

The skill is a plain Python program and runs fine without Claude. The launcher creates its own
venv next to the skill and installs dependencies on first run.

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
| `--threshold` | `0.5` | Minimum Jaccard similarity to treat two slides as the same slide |
| `--output-dir DIR` | temp dir | Where diff files are written |

## Requirements

- **Python >= 3.10** — the only hard requirement. `run.sh` / `run.ps1` find a suitable interpreter
  and bootstrap the venv themselves.
- **[lxml](https://lxml.de/)** — installed automatically into that venv on first run. This first
  run needs network access to reach PyPI; every comparison afterwards is fully offline.

PowerPoint itself is *not* required, and neither is LibreOffice — the skill reads the `.pptx` zip
container directly.

## Repository layout

```
skills/pptx-compare/          the skill itself (this is what you install)
├── SKILL.md                  entry point: strategy selection
├── run.sh / run.ps1          venv bootstrap + dispatch (POSIX / Windows)
├── core/                     extractor / normalizer / mapper / differ
├── slide-compare/            full-detail strategy
└── slide-compare-lite/       fast overview strategy

docs/normalization-dev-guide.md   how to add a normalization rule
AGENTS.md                         orientation for AI coding agents
```

The `skills/<name>/SKILL.md` layout is the conventional one, so skill installers that read a git
repository directly — such as [`skills`](https://github.com/vercel-labs/skills) — pick this up
without any extra configuration.

The normalizer is where most of the domain knowledge lives. If you find a PowerPoint artifact that
still shows up as a false positive, `docs/normalization-dev-guide.md` walks through diagnosing it
and adding a rule.

## Note on language

The split is deliberate: **prompts are Japanese, everything else is English.**

`SKILL.md` and the two strategy documents are prompts — Claude reads them — and they are written
in Japanese, which is why Claude reports its findings in Japanese by default. Everything a person
reads (this README, `AGENTS.md`, `docs/`, all code and comments) is English, as are the JSON
output and the diff files.

To get reports in another language, translate those three prompt files; nothing else needs to
change.

## Contributing

See [DEVELOPMENT.md](DEVELOPMENT.md) for the development setup, the lint command, and a map of the codebase.

Bug reports and feature requests are welcome via GitHub issues.

## License

Apache License 2.0. [LICENSE](LICENSE) holds the license text verbatim; [NOTICE](NOTICE) holds
the copyright attribution, which redistributions must carry along under section 4(d).

`lxml` is installed at runtime into the skill's own venv rather than bundled here, so it is not
redistributed by this project and carries its own license where it is installed.
