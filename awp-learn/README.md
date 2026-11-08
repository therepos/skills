# awp-learn — maintainers' notes

Current: **awp-learn v1.0** (engine: awp-generate v2.1) · libraries in `../awp-lib/`.

`scripts/learn.py` holds only the harvest and merge logic; everything it shares with building (export, token substitution, validation, library discovery) is imported from `awp-generate/scripts/awp.py`, which `release.py` copies into the zip next to `learn.py`. Change shared behaviour in awp-generate, never here.

## Release
`python release.py awp-learn <version>` from the repository root produces `dist/skill-awp-learn-<version>.zip`: this folder + `lib/` (a copy of `awp-lib/`) + `scripts/awp.py`. Before tagging: `python awp-generate/scripts/awp.py check` must pass (it runs learn idempotence through this skill when both are present), then a `learn` on a finished workbook from an earlier release must still produce sensible proposals.

## What the owners do with contributed library files
1. Drop the received `core.json` / `pack-<process>.json` into `awp-lib/` (they carry their own version bump and changelog entry with the reviewer's reasons).
2. `python awp-generate/scripts/awp.py check --deny names.txt` → `CHECK PASSED`.
3. Commit, then release every `awp-*` skill again so each zip carries the new libraries.

## Design decisions
- One merge key: control objective + risk addressed. Refine or replace; never append a near-duplicate.
- Textual matching is advisory; the reviewer's yes/no per item is the control. `apply` enforces only what a script can: the engagement's own words, persons' names, unreviewed numbers and unknown words, dangling references.
- Accepted items carry a fingerprint of the harvested text, so re-running `learn` on the same workbook skips them.
