# awp-generate — running it in corporate Copilot (M365 Copilot / GitHub Copilot, Claude or GPT models)

The model's job is judgement (reading the material, choosing modules, deepening the spec, reviewing learn proposals). The workbook itself is produced by `scripts/awp.py`, so the format never depends on the model. Requirements: a chat that can **unzip a file, run Python 3 with `openpyxl`, and return a file**. Nothing is installed by the user.

## Per session (what the user does)
1. Attach the release zip (`skill-awp-generate-<version>.zip`; for the end-of-engagement harvest, `skill-awp-learn-<version>.zip`) and the engagement documents to the chat.
2. Prompt, for example:
   > Unzip skill-awp-generate, read `SKILL.md` and follow it. Run the selftest. Then `/awp-build` for <client> / <process> / <period> from the attached scope, planning memorandum, prior-year report and SOP.
   With no material: > Unzip skill-awp-generate, read `SKILL.md`, run the selftest, then `/awp-template` with 12 tests.
   End of engagement (awp-learn zip): > Unzip skill-awp-learn, read `SKILL.md`, then `/awp-learn` on the attached finished AWP. Walk me through the proposals one by one.
3. Accept the workbook only after the model reports the verification list (step 7 of `/awp-build`) and the open questions.

## What the model must do first (in every session)
- Unzip into a working folder. The zip unpacks to a single folder (`awp-generate/` or `awp-learn/`); work inside it. `lib/` holds the libraries, `scripts/` the engine, `ref/` the references.
- `python scripts/awp.py selftest` (awp-generate) or `python scripts/learn.py selftest` (awp-learn) → one line starting `SELFTEST PASSED` with the versions. Report the versions to the user. If `openpyxl` is missing: `pip install openpyxl` in the sandbox (it is the only dependency). The selftest and `check` build their files in the system temp folder and delete them.
- Read `SKILL.md`; open the `ref` files when the step needs them (schema for writing the selection, content for writing text, learn for proposals).
- If `python` is not found, try `python3`. If the sandbox blocks `pip`, ask the user to request `openpyxl` — nothing else is needed.

## Sandbox notes
- Formulas (header links, RCM result links, computed columns) calculate when Excel opens the file; no recalculation is needed in the sandbox.
- If the sandbox does not keep files between turns, re-unzip at the start of each turn; the selection and spec JSON delivered with the workbook make any build reproducible.
- Larger models give better content; the format and the checks are identical on every model because they live in the script.

## Prompt pattern for a full build (keeps judgement visible; skip the pauses if you want the result in one go)
1. "Read the material and propose the modules to include, the tokens, and the prior-year points per risk (as a table)." → agree.
2. "Compose, then show me the notes and what you will add to or correct in the spec beyond the library." → agree.
3. "Validate, build, run the verification list and give me the open questions." → deliver.
