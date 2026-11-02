# awp-learn — `/awp-learn`: how a finished AWP improves the libraries without making them messy

## Principle
A library item is identified by **the control objective and the risk it addresses**, not by its wording. If the finished AWP has a better version of an item that is already in the library, the library item is **replaced**. A new item is added only when no library item addresses that objective. Minor variations (a threshold, a role title, a system name, a frequency) are never new items: they become tokens or stay in the engagement spec.

## Step 1 — harvest
`python scripts/learn.py learn "<AWP>.xlsx" proposals.json [--tokens tokens.json] [--lib <folder>]`

- Export is header-driven and tolerant: it finds the `Risk`, `Activity`/`Procedures`, `Question` and `Work done` headers wherever they are, so a workbook edited during fieldwork, or an older-layout file, still harvests (a note is printed for anything it could not find).
- Tokens come from the workbook's hidden `_build` sheet. A workbook built outside the skill has none: pass `--tokens` with the engagement's words (`{"CLIENT": "…", "CLIENT_SHORT": "…", "SYSTEM": "…", "HEAD_OF_DEPARTMENT": "…", …}`).
- Anonymisation: every token value is replaced by its token (whole words only, longest first; number-like values such as "3", "90" or "three" are not substituted because they would corrupt unrelated text); the client's name, abbreviations defined in the text (`("XYZ")`) and company names in work-done rows are replaced by `{CLIENT_SHORT}` or `[REDACTED]`. Then every remaining engagement word, every capitalised word that does not occur in the libraries' own vocabulary, every amount, year, `FY` label, `Observation n`, person title, version number, unit number and "n days/months/%" is listed under **review before accepting**.
- Matching compares the harvested text with the library text after filling the library's tokens with this engagement's words (core defaults < pack defaults < the workbook's tokens), so a workbook built from the current libraries matches exactly. A workbook built from an **older library version** will propose the old wording back as refinements — reject those unless they add substance. `proposals.md` shows the library text as stored, with its tokens, so you can see which parts are variable.
- Matching: controls are matched on control text or activity text, questions on question text, tests on title or objective + steps, risks on risk text, analytics on analytic text. Status `same` (the texts are the same, not reported), `refinement` (a library item with the same objective probably exists; its text is shown), `refinement (extends the library text)` (the harvested text contains the library text and more — usually worth a `replace`), `refinement (uncertain match - check the objective)` (weak textual match — often a `new` item in the right module), `new` (no match; the nearest item is named for orientation only). Items already learned from the same workbook earlier are skipped.
- Harvested: activities with controls, walkthrough questions, tests (fields and steps), RCM risks, DA analytics. Not harvested: DE findings, work-done rows, notes, findings, conclusions — these are fieldwork evidence, not library content.
- Output: `proposals.json` (machine) and `proposals.md` (for the reviewer: every field of each item, the matched library text, the review terms, the tokens used). Both carry the file name and the client's words: engagement working papers, never shipped.

## Step 2 — review with a human (never skip)
For each proposal decide:
- **refinement** → accept as a replacement of the matched item only if it is better in substance (clearer objective, better evidence, a step the library lacked). Wording-only differences: reject.
- **new** → accept only if the objective is not already covered elsewhere in the module list; choose the module (or name a new one) and the library (`core` for anything that applies to every process; `pack:<process>` otherwise).
- Before accepting, edit the proposal's text so that nothing under *review before accepting* remains unless it is plainly generic (then set `"redaction_checked": true`). Replace specifics with tokens or general wording ("in accordance with the Delegation of Authority", "within the committed timeline", "the month of the operational change").
- Prior-year specifics never enter a library; a generic "why we ask" ("a frequent prior-year finding is followed up here") is the ceiling.

## Step 3 — decisions file
```json
{
  "P001": {"accept": true, "library": "core", "module": "mod.policies_doa", "replace": "ctrl.sop_approval",
           "edits": {"control": "The SOP is reviewed and approved by the {HEAD_OF_DEPARTMENT} and countersigned by the {HEAD_OF_FINANCE} where it affects billing …"}},
  "P002": {"accept": true, "library": "pack:commercial-leasing", "module": "mod.cl_deposit_refund", "id": "ctrl.cl_refund_ageing", "risk": "risk.cl_deposit_refund",
           "edits": {"control": "… any refund outstanding beyond the escalation threshold in the SOP is escalated to the {HEAD_OF_DEPARTMENT}.", "document": "Refund ageing reports for the sampled months"}},
  "P003": {"accept": true, "library": "pack:commercial-leasing", "module": "mod.cl_adjustments", "id": "q.cl_rebate_programme", "control": "ctrl.cl_credit_note",
           "edits": {"question": "How are rent rebates granted under a programme approved by the appropriate authority recorded in {SYSTEM} …"}},
  "P004": {"accept": true, "library": "pack:commercial-leasing", "module": "mod.cl_incentives", "module_title": "Lease incentives", "sheet": "tenant_lease", "stage": "cycle", "order": 25,
           "token_defaults": {"{INCENTIVE_CAP}": "the cap in the leasing policy"}, "redaction_checked": true},
  "P005": {"accept": false}
}
```
- `library`: `core` or `pack:<name>`; defaults to the matched library. `module`: an existing module id, or a new id with `module_title`, `sheet` and `stage` (and `order`).
- `edits`: the final wording of any harvested field (`control`, `activity`, `question`, `why`, `steps`, …). This is where the generalisation is done; the original proposal text stays untouched.
- `replace`: id of the library item to overwrite (refinement). Without it, the item is appended with `id` (or a generated one).
- `risk`: for a new control or test, the risk id to link it to (otherwise it will not appear in the RCM — `validate` warns). `control`: for a question, the control id it tests. For a new risk, `controls` and `tests` list the ids it maps to; for a new analytic, `tests`.
- `token_defaults`: declares a new token (with its default) in the library when a threshold or name should become a token.
- `redaction_checked`: a list of the review terms that are plainly generic (`["45 days", "Board"]`), or `true` to accept all of them. The engagement's own words (token values that differ from the library defaults, the client name with or without its corporate suffix, the denylist), a person's title and name (`Mr Tan`) and `[REDACTED]` are refused regardless — in the text, in `edits`, and in `id`, `module_title` and `token_defaults`.
- `reason`: free text; it is written into the library's changelog entry.
- A `refinement` that matched the wrong item is accepted as new by omitting `replace` and giving `module`, `id` and (for a control or test) `risk`.
- A `module` that does not exist is created only when `module_title`, `sheet` (an existing sheet id) and `stage` are all given; otherwise `apply` stops, so a typo cannot create a stray module.

## Step 4 — apply
`python scripts/learn.py apply proposals.json decisions.json [--lib <folder>]`
- Re-scans the final text of each accepted proposal (harvested fields with `edits` applied) and the decision's own ids, titles and token defaults. It refuses the engagement's own words and persons' names outright, and anything on the review list that the decision has not accepted.
- What it cannot catch: a lower-case proper noun ("the asset-enhancement programme"), a bare surname, a client fact written in plain words. The reviewer is the last line of defence — read every accepted text once as a stranger would.
- Writes the updated library files **in place** in `lib/` (the previous versions are in the release zip and in the repository; copy the folder first if you want a local backup).
- Writes the item, bumps the library `version`, sets `updated`, appends a changelog entry inside the file (source is recorded as `/awp-learn` with each decision's `reason`, never the file name), re-validates the library (duplicate ids, unresolved references, near-duplicate controls, undeclared tokens) and refuses to save on an error.
- Output: the updated `lib/*.json` only. Hand them to the skill owners with the changelog line; they run the acceptance tests and publish the next release zip. Layer 1 is untouched.

## Idempotence
Running `learn` on a workbook that was built from the current libraries and not edited yields zero proposals. A build that was deepened through `overrides` or `extra_modules` will, by design, propose those additions back — that is how engagement content reaches a library. Running it on the same finished AWP twice yields the same proposals. After `apply`, each accepted item carries the fingerprint of the harvested text it came from, so a later `learn` on the same workbook skips it even though the library wording was generalised (`already learned` in the summary).
