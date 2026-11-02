# skills

Claude skills, one folder per skill (`<name>/SKILL.md` plus `scripts/`).

To use a skill in another repo, copy its folder into that repo's `.claude/skills/<name>/`.
This repo is the source of truth; update here first, then re-copy.

## skill-awp — internal-audit work program builder

`skill-awp/` uses a flat layout by naming convention instead of `SKILL.md` + `scripts/`: the instructions file is `skill-awp/skill-awp.md` (a strict loader can rename it to `SKILL.md`), the engine is `skill-awp/skill-awp.py`, references are `skill-awp-ref-*.md`, content libraries are `skill-awp-lib-*.json`. The standard behind the layout is documented in `skill-awp/skill-awp-readme.md`. Users of corporate Copilot get a release zip of the folder; `python skill-awp/skill-awp.py check` must print `CHECK PASSED` before a release.
