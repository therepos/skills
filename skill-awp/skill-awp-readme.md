# skill-awp — maintainers' notes (ships with the release; engagement users can ignore it)

Current: **skill-awp v2.0** · core v1 · pack commercial-leasing v1. Versions live inside the files (`VERSION` in the script, `"version"` in each library); file names never change. History is GitHub; a release is the zip of this folder.

## Layout (flat, by naming convention — see `../skills-readme.md`)
```
skill-awp.md                              instructions and commands (Layer 1)
skill-awp.py                              the engine (Layer 1)
skill-awp-ref-layout.md                   sheet layout and formatting the engine enforces
skill-awp-ref-schema.md                   selection / spec / library JSON and the token table
skill-awp-ref-content.md                  writing rules with generic examples
skill-awp-ref-learn.md                    /awp-learn rules and the decisions file
skill-awp-ref-copilot.md                  running it in corporate Copilot
skill-awp-ref-selection-example.json      example selection (fictitious client) - also the reference build for /awp-release
skill-awp-ref-example.xlsx                the workbook built from it (what every build looks like)
skill-awp-lib-core.json                   universal modules (Layer 2)
skill-awp-lib-pack-commercial-leasing.json process pack (Layer 2); add skill-awp-lib-pack-<process>.json for new processes
skill-awp-readme.md                       this file
skill-awp-changelog.md                    release history
```

## `/awp-release` — publishing a new version (owners)
1. Pull the proposed changes: engine changes as a PR; library changes as the updated `skill-awp-lib-*.json` files produced by `apply` (their internal changelog says what changed and when).
2. Acceptance tests: `python skill-awp.py check --deny names.txt` must print `CHECK PASSED` (`names.txt` = one client, system or person name per line that must never appear in a library; keep it outside the repository). It runs the selftest (repeat it once in a clean environment with only `openpyxl`), validates every library, composes and builds the reference selection, round-trips it through `export`, runs `learn` on it (must propose nothing), scans the build for unresolved tokens or placeholders and the libraries for amounts, observation numbers and the denied names. Then rebuild `skill-awp-ref-example.xlsx` from the reference selection and `diff` it against the previous release's copy: every differing cell must be explained by a library change.
3. Bump `VERSION` in `skill-awp.py` only if the engine changed; library versions are already bumped by `apply`. Add a line to `skill-awp-changelog.md`.
4. Zip the folder as `skill-awp-<version>.zip` (and the extension-less 7z variant if the corporate chat strips extensions). Users download the zip; GitHub keeps the raw files and history.

## Design decisions (so they are not reopened by accident)
- **Structure in code, content in libraries, specifics in tokens.** Users add depth through libraries and the engagement spec; they never alter the structure.
- **Id-based modules with composition in code.** A-numbers, control numbers and cross-references are assigned at compose time, so a library never breaks a reference.
- **One merge key for learning: control objective + risk addressed.** Refine or replace; no near-duplicate appends. The validator flags near-duplicate controls.
- **Minimal controls.** No manifests or signatures; version stamp in every workbook, versions inside files, GitHub history. Anyone with the folder can edit Layer 1; the stamp and the changelog make that visible, which is the proportionate control.
- **Users see Excel only.** The selection and spec JSON travel with the engagement but need never be opened; the hidden `_build` sheet is six rows.
- **DA sheet is `A400`** for new builds; `DA` is accepted on export for legacy workbooks.
- **Library `denylist`** (a list of words in a library file) is optional: owners may list names that must never appear; `validate` then fails on them. The release check does the same with `--deny names.txt`, which is the usual place for the list.

## Adding a process pack
Engagement teams produce draft packs (`"version": 0`, built with `--lib`, stamped "(draft)") when no pack fits their process; `/awp-build` step 2 tells them how. To publish one: validate it, read every module for client words (`check --deny`), set `token_defaults` for the process's `{TRANSACTION_TYPES}` and `{PROCESS_ACTIVITIES}`, add a `sheet_map` if a core module belongs on one of the pack's sheets, set `"version": 1`, compose it with core against a test selection, build, `verify`, review the workbook, add the pack name to the changelog and ship it in the next release.
