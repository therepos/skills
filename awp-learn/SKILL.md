---
name: awp-learn
description: Harvest a finished internal-audit work program (AWP) workbook into the shared content libraries - anonymised proposals, a human yes/no per item, then apply - so the next engagement starts from what this one learned. Use at the end of an engagement when a user says "update the skill with what we did", "add what we learned", "save this for the next team" or mentions /awp-learn.
---

# awp-learn — harvest a finished AWP into the shared libraries

**Use this skill when** an engagement has ended and the user wants what the team did in the workpaper to benefit the next team. Building a new AWP is the **awp-generate** skill, not this one.

**What it produces**: the updated library JSON file(s) — `lib/core.json` and/or `lib/pack-<process>.json` — carrying the accepted items in generic form, plus a short plain-language note. Nothing else changes. The user sends the file(s) to the skill owners, who publish the next release of every `awp-*` skill. A new zip is never the deliverable of an engagement session.

## The user says… → what you do
| The user says something like | Do |
|---|---|
| "update the skill / library with what we did", "add what we learned", "save this for the next team" | the `/awp-learn` procedure below |
| "build / prepare the AWP", "give me the template" | not this skill: point them to **awp-generate** (`skill-awp-generate-<version>.zip`) |
| "change the format / the script / add a column" | decline: Layer 1 belongs to the skill owners; tell them what to ask for |

Never edit `scripts/*.py`, `SKILL.md`, the `ref` files or a library by hand: libraries change only through `apply`.

## Libraries hold no client data
No client, tenant, vendor, system, person, amount, date, site or prior-year observation ever enters a library. Specifics become tokens (`{SYSTEM}`, `{HEAD_OF_DEPARTMENT}`, `{REFUND_DAYS}` …) or general wording ("in accordance with the Delegation of Authority", "within the committed timeline"). `apply` refuses the engagement's own words and persons' names outright and stops on unreviewed numbers, dates and unknown words — but it cannot recognise a client fact written in plain lower-case words, so **the reviewer reads every accepted text once as a stranger would**. Do not work around a refusal.

## `/awp-learn` — the procedure
Work inside the folder that contains `SKILL.md` (the zip unpacks to `awp-learn/`; libraries are in `lib/`, the engine in `scripts/`). Run `python scripts/learn.py selftest` first (`python3` if `python` is not found); it must print `SELFTEST PASSED` with the library versions. Needs Python 3.6+ and `openpyxl`, nothing else.

1. **Harvest**: `python scripts/learn.py learn "<finished AWP>.xlsx" proposals.json` (add `--tokens tokens.json` if the workbook has no hidden `_build` sheet, was built for another client, or the names have changed — `--tokens` overrides what `_build` says; the tokens used are printed at the top of `proposals.md`). The script exports the visible sheets, replaces the engagement's words with tokens, compares every control, question, test, risk and analytic with the libraries, and writes `proposals.json` + `proposals.md` (full text of each item, the matched library text with its tokens, and the words to review). Fieldwork results — DE findings, work-done rows, notes, findings, conclusions — are engagement evidence and are never harvested.
2. **Review with the user** (or, if they are not available, decide as a careful reviewer would and record the reason): for each item in `proposals.md`, `refinement` means a library item with the same objective exists and its text is shown — accept with `replace` only if the harvested version is better in substance **and the matched item really has the same objective** (the match is textual; "uncertain match" means check it; when it is wrong, accept as `new` in the right module instead); `refinement (extends the library text)` is usually worth a replace; `new` means no match. Never add a near-duplicate. A minor variation (a threshold, a role name, a system) becomes a token (`token_defaults` in the decision) or stays in the engagement. Words under *review before accepting* may identify the engagement — generalise or delete them through `edits`. A workbook built with an older library version proposes its old wording back; reject those unless they add substance.
3. **Decide**: write `decisions.json` (format and a worked example in `ref/learn.md`; `edits` carries the final wording, `reason` the rationale, `risk` links a new control or test to its RCM row, `control` links a question to its control, `redaction_checked` lists the review terms that are plainly generic).
4. **Apply**: `python scripts/learn.py apply proposals.json decisions.json`. The script re-scans the final text and the decision's ids and titles, merges, bumps the library version, appends the changelog inside the file, re-validates, and writes the library file(s) in place in `lib/`.
5. **Deliver** the changed file(s) from `lib/` (only the ones `apply` reported) and a plain-language note: what the next team gets, what was deliberately not kept and why, and the changelog line for the owners. `proposals.json`/`.md` and `decisions.json` carry the client's words and the file name: they stay with the engagement, never with the owners.

## Files
- `SKILL.md` — this file. `README.md` — maintainers' notes.
- `scripts/learn.py` — learn / apply / selftest. `scripts/awp.py` — the shared engine from awp-generate (export, tokens, validation), bundled into the release zip.
- `lib/core.json`, `lib/pack-*.json` — the libraries as released (in the repository: `../awp-lib/`, shared by every `awp-*` skill).
- `ref/learn.md` — matching rules, anonymisation, the decisions file, what `apply` refuses, idempotence.
