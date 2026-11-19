# awp-update — maintainers' notes

Current: **awp-update v1.1** (engine: awp-generate v2.1). No libraries: the skill reads and writes an engagement workbook only.

`scripts/update.py` holds the read / plan / apply logic. The house format is not duplicated here: changed sheets are rebuilt with `build_rcm`, `build_walkthrough` and `build_test` from `awp-generate/scripts/awp.py`, which `release.py` copies into the zip next to `update.py`, and the workbook is read with its `export_spec`. Change the format in awp-generate, never here.

## Release
`python release.py awp-update <version>` from the repository root → `dist/skill-awp-update-<version>.zip` and `dist/copilot-awp-update-<version>.zip` (this folder + `scripts/awp.py`; no `lib/`). Before tagging: `python awp-update/scripts/update.py selftest`, then one full run on `awp-generate/ref/example.xlsx` saved once by a spreadsheet application (so formulas carry cached values), with staff responses typed in, using the worked example in `ref/update.md`: `apply` must print `VERIFY PASSED`, and every sheet without changes must be cell-for-cell identical to the input. Release again whenever awp-generate's engine changes the layout of the 100, 200 or 300 sheets.

## Design decisions (so they are not reopened by accident)
- **The model writes a change list, the script writes the workbook.** Cell-by-cell editing by a model is what times out and what drifts from the format. An update file is small, reviewable and re-runnable.
- **Rebuild only the sheets that change, inside the user's workbook.** A full rebuild from an exported spec would drop anything the team added outside the house tables on any sheet; cell-level patching would need a second implementation of the layout. Rebuilding a changed sheet from its exported content with the engine's own builders keeps one implementation of the format and leaves every other sheet untouched.
- **Nothing is lost silently.** Before writing, each sheet to be rebuilt is round-tripped (export → build) and compared with the original: text outside the tables, comments, pictures, charts, and embedded files anywhere in the workbook are listed and stop the update unless the user accepts with `--force`.
- **Staff responses are protected in code**, not by instruction: overwriting needs an explicit `mode`.
- **Every change needs a source and a basis**, and all of it goes to an update log beside the workbook. The log is a separate file so the workbook's structure stays the house structure.
- **Row numbers in update files are the original workbook's.** All files are applied in one `apply`; renumbering happens once at the end. Control numbers never change; new controls take the next free number on their sheet.
- **Rows are rewritten, rarely removed.** A planned control the client lacks is a design gap to record, not a row to delete; `remove_activity` exists for activities that do not exist at the client and cleans the references it leaves behind.
- **The RCM follows the test.** When a test's steps change, the RCM row that quoted the full steps is updated with it; rows that quote selected steps are flagged by the engine's validator.
- **Expected becomes actual, enforced in code.** Planning labels ("Expected: …", "Expected process") are detected by pattern: `show` tags them, the still-open list reports them, and `check` refuses new text that keeps one. The pattern matches labels, not the ordinary word ("arrears are expected to be settled" in a library background is not a label). Rows the sources do not cover keep their label on purpose.
- **Out of scope for now:** recording test work (300 series) and conclusions; harvesting (awp-learn).
