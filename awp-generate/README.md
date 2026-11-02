# awp-generate — maintainers' notes (ships with the release; engagement users can ignore it)

Current: **awp-generate v2.1** · core v1 · pack commercial-leasing v1. Versions live inside the files (`VERSION` in the script, `"version"` in each library); file names never change. History is GitHub; a release is `python release.py awp-generate <version>` from the repository root, which zips this folder with a copy of `../awp-lib/` as `lib/`.

## Layout (repository)
```
awp-lib/                         shared content libraries (Layer 2) - read by every awp-* skill
  core.json                      universal modules
  pack-commercial-leasing.json   process pack; add pack-<process>.json for new processes
awp-generate/                    this skill (Layer 1)
  SKILL.md                       instructions and commands - the entry point
  README.md, CHANGELOG.md        these notes, release history
  scripts/awp.py                 the engine
  ref/layout.md                  sheet layout and formatting the engine enforces
  ref/schema.md                  selection / spec / library JSON, token table, validator rules
  ref/content.md                 writing rules with generic examples
  ref/copilot.md                 running it in corporate Copilot
  ref/selection-example.json     example selection (fictitious client) - also the reference build for the release check
  ref/example.xlsx               the workbook built from it (what every build looks like)
awp-learn/                       sister skill: harvests finished AWPs back into awp-lib (imports scripts/awp.py)
release.py                       builds the self-contained zip of a skill (adds lib/, and awp.py for awp-learn)
```
In a release zip the same files sit under `awp-generate/` with `lib/` beside `scripts/`; the script looks for `lib/` first, `../awp-lib/` second.

## `/awp-release` — publishing a new version (owners)
1. Pull the proposed changes: engine changes as a PR; library changes as the updated `core.json` / `pack-<process>.json` files produced by awp-learn's `apply` (their internal changelog says what changed and when).
2. Acceptance tests: `python awp-generate/scripts/awp.py check --deny names.txt` (from the repository root, with awp-learn beside it) must print `CHECK PASSED` (`names.txt` = one client, system or person name per line that must never appear in a library; keep it outside the repository). It runs the selftest (repeat it once in a clean environment with only `openpyxl`), validates every library, composes and builds the reference selection, round-trips it through `export`, runs `learn` on it (must propose nothing), scans the build for unresolved tokens or placeholders and the libraries for amounts, observation numbers and the denied names. Then rebuild `ref/example.xlsx` from the reference selection and `diff` it against the previous release's copy: every differing cell must be explained by a library change.
3. Bump `VERSION` in `scripts/awp.py` only if the engine changed; library versions are already bumped by `apply`. Add a line to `CHANGELOG.md`.
4. `python release.py awp-generate <version>` → `dist/skill-awp-generate-<version>.zip` (and an extension-less 7z copy for chats that strip extensions); attach it to a GitHub release tagged `awp-generate-<version>`. Users download the zip; GitHub keeps the raw files and history. When a library changed, release awp-learn too.

## Design decisions (so they are not reopened by accident)
- **Structure in code, content in libraries, specifics in tokens.** Users add depth through libraries and the engagement spec; they never alter the structure.
- **Id-based modules with composition in code.** A-numbers, control numbers and cross-references are assigned at compose time, so a library never breaks a reference.
- **One merge key for learning: control objective + risk addressed.** Refine or replace; no near-duplicate appends. The validator flags near-duplicate controls.
- **Minimal controls.** No manifests or signatures; version stamp in every workbook, versions inside files, GitHub history. Anyone with the folder can edit Layer 1; the stamp and the changelog make that visible, which is the proportionate control.
- **Users see Excel only.** The selection and spec JSON travel with the engagement but need never be opened; the hidden `_build` sheet is six rows.
- **DA sheet is `A400`** for new builds; `DA` is accepted on export for legacy workbooks.
- **Library `denylist`** (a list of words in a library file) is optional: owners may list names that must never appear; `validate` then fails on them. The release check does the same with `--deny names.txt`, which is the usual place for the list.
- **Libraries are shared and bundled.** One copy in `awp-lib/` in the repository; every `awp-*` release zip carries a snapshot as `lib/`. Users never need the repository.

## Adding a process pack
Engagement teams produce draft packs (`"version": 0`, built with `--lib`, stamped "(draft)") when no pack fits their process; `/awp-build` step 2 tells them how. To publish one: validate it, read every module for client words (`check --deny`), set `token_defaults` for the process's `{TRANSACTION_TYPES}` and `{PROCESS_ACTIVITIES}`, add a `sheet_map` if a core module belongs on one of the pack's sheets, set `"version": 1`, save it as `awp-lib/pack-<process>.json`, compose it with core against a test selection, build, `verify`, review the workbook, add the pack name to the changelog and release every `awp-*` skill.
