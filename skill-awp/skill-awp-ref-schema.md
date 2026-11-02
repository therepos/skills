# skill-awp — JSON reference (selection, spec, library)

All files are UTF-8 JSON. Ids are lower-case dotted strings and must be unique within a library (`mod.`, `ctrl.`, `q.`, `risk.`, `test.`, `da.`). References inside libraries use ids only; the script assigns A-numbers.

## 1. Selection (input to `compose`) — one per engagement
```json
{
  "client": "Client Pte Ltd", "process": "Commercial Leasing", "period": "1 April 2026 to 31 March 2027",
  "prefix": "A", "prepared_by": "", "reviewed_by": "",
  "packs": ["commercial-leasing"],
  "modules": [], "exclude": [],
  "sheet_map": {},
  "tokens": { "CLIENT_SHORT": "CPL", "SYSTEM": "PropertySys", "ERP": "the general ledger system", "HEAD_OF_DEPARTMENT": "Head of Leasing" },
  "prior_year": { "risk.cl_deposit_refund": "In the prior year we raised Observation 2.3 (Low) …" },
  "background_additions": { "control_environment": "In the period the company moved to …" },
  "overrides": { "test.cl_closures": {"basis": "All 22 closures in the east wing."}, "q.cl_exit": {"why": "…"}, "ctrl.cl_refund_approval": {"text": "…", "activity": "…"}, "risk.access": {"risk": "…"}, "mod.cl_sourcing": {"scope_areas": ["…"]} },
  "extra_modules": [ { "id": "mod.eng_eastwing", "title": "East-wing relocations", "stage": "cycle", "order": 35, "scope_areas": ["…"], "walkthrough": {"sheet": "tenant_lease", "section_title": "…", "activities": [], "questions": []}, "risks": [], "tests": [], "analytics": [] } ],
  "extra_sheets": [],
  "extra": {}
}
```
- `packs`: names after `skill-awp-lib-pack-`. Core is always loaded. Order matters only when two libraries define the same sheet id (later wins).
- `modules` (ids to include; empty = all) and `exclude` narrow the scope. `sheet_map` moves a module to another walkthrough sheet (`{"mod.master_file": "tenant_lease"}`); packs ship their own `sheet_map`.
- `prior_year`: keyed by risk id (`list` prints them). The text is appended to the RCM "Finding for Design Effectiveness" cell of that risk, to the "why we ask" of the first question on each of the risk's controls, and "This test also follows up on prior-year Observation n." to the objective of the risk's tests (the observation number is read from the text).
- `background_additions`: keyed by walkthrough sheet id (`list` prints them; a pack's `sheet_map` may move a core module onto a pack sheet). The paragraph is appended to the sheet's background.
- `overrides`: keyed by any library id (`list` prints them): the fields given replace the library's before composition. For a control id, `activity`, `performed_by`, `frequency` and `document` go to its activity row and everything else (`text`, `owner`, `classification`, `type`) to the control. Tokens may be used in the new text.
- `extra_modules` / `extra_sheets`: engagement-specific modules and sheets in library format (section 3), composed, numbered and cross-referenced like any pack module; they may reference core and pack ids (e.g. a test id in a risk). They never go into a library unless `/awp-learn` proposes them later.
- `extra`: replaces a whole top-level spec key verbatim after composition; for targeted changes use `overrides` or edit the spec itself.
- Unknown ids in any of these produce a `note` from `compose`.

### Tokens
Resolution order: core `token_defaults` < pack `token_defaults` < derived from the selection < `tokens`. Keys may be written with or without braces. Derived automatically: `{CLIENT}` = client, `{CLIENT_SHORT}` = client unless given, `{PROCESS}` = process in lower case, `{DEPARTMENT}` = process, `{PERIOD}` = period, `{HEAD_OF_DEPARTMENT}` = "Head of <department>", `{MANAGER}` = "Manager (<department>)", `{STAFF}` = "<department> staff". `python skill-awp.py list` prints every token with its default; `compose` prints the ones left at their default.

| Token | Meaning | Core default |
|---|---|---|
| `{CLIENT}` / `{CLIENT_SHORT}` | legal name / short name | from selection |
| `{PROCESS}` / `{DEPARTMENT}` | process (lower case) / department name | from selection |
| `{HEAD_OF_DEPARTMENT}` / `{MANAGER}` / `{STAFF}` | head, manager and staff-level role of the department | derived |
| `{FINANCE}` / `{FINANCE_PREPARER}` / `{FINANCE_REVIEWER}` / `{HEAD_OF_FINANCE}` | finance function and its three roles | Finance / Finance preparer / Finance reviewer / Head of Finance |
| `{CEO}` / `{IT}` | chief executive / IT function | CEO / IT |
| `{SYSTEM}` / `{ERP}` | operational system / general ledger system | the system / the general ledger system |
| `{COUNTERPARTY}` / `{COUNTERPARTIES}` / `{COUNTERPARTY_TITLE}` | tenant, vendor, customer … | counterparty / counterparties / Counterparty |
| `{AGREEMENT}` / `{MASTER_FILE}` | contract name / master file name | contract / master file |
| `{CURRENCY}` / `{RECON_BALANCES}` / `{INTERFACE_FREQUENCY}` | currency label / balances reconciled between the systems / how often {SYSTEM} posts to {ERP} ("every month", "nightly") | S$ / the balances recorded in both systems / every month |
| `{TRANSACTION_TYPES}` / `{PROCESS_ACTIVITIES}` | what the Delegation of Authority covers / the activities the segregation-of-duties question walks through | generic; the pack supplies the process's list |
| `{SUBLEDGER}` / `{VARIANCE_SUBJECT}` | what bank movements are reconciled to / what the monthly variance analysis covers | the sub-ledger / the amounts recorded, by type and by {COUNTERPARTY} |
| `{MASTER_FILE_KEYER}` / `{MASTER_FILE_APPROVER}` | who keys and who approves master-file records | {STAFF} / {HEAD_OF_DEPARTMENT} (a token's default may contain other tokens) |
| pack tokens | e.g. leasing: `{FACILITIES}`, `{DEPOSIT_MONTHS}`, `{DEPOSIT_MONTHS_N}`, `{REFUND_DAYS}`, `{REFUND_FROM}` ("the lease end date" / "handover") | declared in the pack's `token_defaults` |

A library may use any token it declares in its own `token_defaults`; `validate` flags undeclared ones.

## 2. Spec (output of `compose`, input to `build`; also what `export` returns)
```json
{
  "prefix": "A", "client": "", "process": "", "period": "", "prepared_by": "", "reviewed_by": "",
  "signoff": true, "stamp": true, "link_results": true,
  "library_versions": {"core": 1, "pack commercial-leasing": 1}, "module_ids": ["mod.policies_doa", "…"], "tokens": {"{SYSTEM}": "PropertySys"},
  "scope": [{"no": 1, "key_process": "Overall control environment", "areas": ["Check that …"], "covered": "Yes"}],
  "rcm": {"sections": [{"title": "1. Overall control environment", "rows": [
      {"sn": 1, "id": "risk.policies", "sub_process": "Policies and procedures", "risk": "… This may result in …",
       "control_no": "A201-1\nA201-2", "de_finding": "To be concluded after the walkthrough.", "audit_strategy": "Test 1: …\nCheck that:\n1. …",
       "testing_result": "", "wp_ref": "A301"}]}]},
  "walkthroughs": [{"ref": "A201", "id": "control_environment", "title": "…", "index_text": "…", "background": "…", "questions_title": null,
      "sections": [{"title": "1.1 …", "module": "mod.policies_doa", "activities": [
          {"no": 1, "activity": "…", "performed_by": "…", "frequency": "Monthly", "control_ref": "A201-1", "control": "…", "owner": "…",
           "classification": "Preventive", "type": "Manual", "de_finding": "", "document": "…"},
          {"no": 2, "refer_to": "Please refer to <A302> for the segregation of duties matrix."}]}],
      "questions": [{"no": 1, "id": "q.sop_version", "question": "…", "why": "…", "documents": "…", "interviewee": "…", "response": "", "control_ref": "A201-1"}]}],
  "tests": [{"ref": "A301", "id": "test.policies_doa", "no": 1, "title": "…", "objective": "…", "basis": "…", "population": "…", "sample_size": "…", "source": "…",
      "steps": "Test 1: …\nCheck that:\n1. …\n2. …", "detail_title": "Lease details", "detail_columns": ["Tenant", "Unit", "…"],
      "detail_formulas": {"Days from lease end to refund": "=IF(AND(ISNUMBER(E{r}),ISNUMBER(H{r})),H{r}-E{r},\"\")"},
      "n_steps": 6, "n_rows": 10, "rows": [{"values": ["…"], "steps": ["✓"], "remarks": ""}], "notes": "Nil", "findings": "Nil", "conclusion": "To be concluded on completion of fieldwork."}],
  "da": {"ref": "A400", "analytics": [{"sn": 1, "id": "da.user_access", "analytic": "…", "source": "…", "identifies": "…", "feeds": "A303", "result": "", "status": "Planned"}], "note": null}
}
```
- Refs: walkthrough sheets `{P}201…`, tests `{P}301…`, controls `{sheet}-{n}`, DA sheet `da.ref` (default `{P}400`; legacy workbooks may use `DA`). `normalise()` fills missing refs in order.
- `link_results: true` puts a formula in the RCM "Testing Result" that shows the Conclusion cell of each WP-ref test sheet (several are joined with line breaks).
- `detail_formulas`: per detail column, `{r}` is the row number; `n_steps` must equal the number of numbered audit steps.
- `stamp: false` suppresses the build stamp and `_build` sheet (not recommended).

## 3. Library (`skill-awp-lib-core.json`, `skill-awp-lib-pack-<process>.json`)
```json
{
  "library": "core" | "pack", "process": null | "commercial-leasing", "version": 1, "updated": "2026-10-28", "description": "…",
  "token_defaults": {"{SYSTEM}": "the system"}, "sheet_map": {"mod.master_file": "tenant_lease"}, "denylist": [],
  "sheets": [{"id": "control_environment", "title": "Overall control environment", "index_text": "…", "background": "… {SYSTEM} …"}],
  "modules": [{
    "id": "mod.system_access", "title": "System access rights", "stage": "control_environment" | "cycle" | "closing", "order": 30,
    "scope_areas": ["Check that access rights to {SYSTEM} …"],
    "walkthrough": {"sheet": "control_environment", "section_title": "System access rights",
      "activities": [{"activity": "…", "performed_by": "…", "frequency": "…", "document": "…",
                      "control": {"id": "ctrl.access_joiner", "text": "…", "owner": "…", "classification": "Preventive", "type": "Manual", "document": "…"}},
                     {"refer_to": "Please refer to <test.sod> for the segregation of duties matrix."}],
      "questions": [{"id": "q.access_users", "question": "…", "why": "…", "documents": "…", "interviewee": "{IT}", "control": "ctrl.access_review"}]},
    "risks": [{"id": "risk.access", "sub_process": "System access rights", "risk": "… Unauthorised transactions may therefore not be detected.",
               "controls": ["ctrl.access_joiner", "ctrl.access_review"], "tests": ["test.access"], "strategy_steps": [5, 6]}],
    "tests": [{"id": "test.access", "title": "…", "objective": "…", "basis": "…", "population": "…", "sample_size": "…", "source": "…",
               "steps_intro": "Check that:", "steps": ["…", "…"], "detail_title": "…", "detail_columns": ["…"], "detail_formulas": {}, "n_rows": 15}],
    "analytics": [{"id": "da.user_access", "analytic": "…", "source": "…", "identifies": "…", "tests": ["test.access"]}]
  }],
  "changelog": [{"version": 1, "date": "2026-10-28", "source": "/awp-learn", "changes": ["…"]}]
}
```
How `compose` turns modules into a workbook:
- Walkthrough sheets appear in the order of the first module that uses them (modules sorted by stage, then `order`); every module becomes one numbered sub-section (`2.3 …`) on its sheet; its questions follow in the same order.
- Tests are numbered in sheet-then-module order; controls `{sheet}-{n}` in order of appearance; `<test.id>` inside `refer_to` text becomes the sheet ref.
- RCM: one section per walkthrough sheet, one row per risk. The first risk (in module order) that uses a test carries the test's full steps as its audit strategy; later risks say "Refer to Test n." (`strategy_steps` lists the step numbers to quote) — put the main risk for a test first in its module. A module with questions but no activities (e.g. `mod.general`) adds its questions without a sub-section row.
- A library with `"version": 0` is a draft; the build stamp says so.
- The walkthrough's Performed by / Frequency / Owner / Classification / Type cells are capitalised by `compose`, so one lower-case token (`{PAYMENT_RUN_FREQUENCY}` = "weekly") serves both the prose and the column.
- Core sentences that use the derived role tokens: `{MANAGER}` prepares the SOP (mod.policies_doa); `{STAFF}` keys master records (mod.master_file, via `{MASTER_FILE_KEYER}`) and prepares refund forms (leasing pack); `{HEAD_OF_DEPARTMENT}` approves access requests, master records and the SOP. If the engagement has no such role, set the token to the role that does it, or override the sentence.

### What `validate` checks in a library (so a pack validates first time)
- Ids unique across the library; every `risks[].controls`, `risks[].tests`, `analytics[].tests` and `questions[].control` id exists (core's tests and controls count as existing for a pack); every control linked to at least one risk (warning).
- Each risk states a consequence: the text contains one of "may result", "resulting in", "so that", "leading to", "which would", "this may", "could result", "exposing", "leaving", "without this being detected", "not detected", "misstated", "unsecured", "loss", "disputes", "unauthorised", "incorrect", "late", "incomplete" (warning).
- Tokens: the standard list in section 1, core's `token_defaults`, and the library's own `token_defaults`; anything else is flagged as non-standard (warning).
- Near-duplicate controls (text similarity above 0.9) are flagged (warning); a library `denylist` entry found in the text is an error.
- A spec is checked for unresolved tokens and cross-references, duplicate refs, placeholders, RCM strategies that quote a test's steps differently from the test, and minimum depth (≥ 3 numbered steps per test, ≥ 2 controls and ≥ 5 questions per walkthrough sheet).
- Scope: one key process per walkthrough sheet; areas = the `scope_areas` of its modules.
- A pack may reference a core test id in `risks[].tests` (e.g. `test.reconciliations`, `test.variance_analysis`).
