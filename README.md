# skills

Claude skills, one folder per skill (`<name>/SKILL.md` plus `scripts/`).

To use a skill in another repo, copy its folder into that repo's `.claude/skills/<name>/`.
This repo is the source of truth; update here first, then re-copy.

## Layout of a skill

```
<name>/
├── SKILL.md        entry point: what the skill does, the commands, the rules (a model reads this first)
├── README.md       maintainers' notes (release steps, design decisions), CHANGELOG.md
├── scripts/        the engine - deterministic work (format, numbering, checks) lives in code, never in the prompt
├── ref/            references the model opens only when a step needs them
└── lib/            (release zip only) data the skill reads; in the repository it may be a shared folder instead
```

Rule of thumb behind every skill here: anything that must stay constant or consistent is code; the model supplies judgement.

## Releases

Users never read this repository; they get a self-contained zip. `python release.py <skill> <version>` builds `dist/skill-<name>-<version>.zip` (copying in whatever the skill needs to run alone), which is attached to a GitHub release tagged `<name>-<version>`. The zip name keeps the `skill-` prefix so it is recognisable on a user's machine; the folder inside is just `<name>/`.

## The awp family — internal-audit work programs

| Folder | Role |
|---|---|
| `awp-lib/` | shared content libraries (`core.json`, `pack-<process>.json`); not a skill, no SKILL.md; bundled into every `awp-*` zip as `lib/` |
| `awp-generate/` | build the AWP workbook from engagement material and the libraries (template, build, verify) |
| `awp-learn/` | end of engagement: harvest the finished AWP back into the libraries with client-data guards |

Further `awp-*` skills (a review skill, say) read the same `awp-lib/`.
