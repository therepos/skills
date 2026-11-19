# awp-update — the update file (what you write; `scripts/update.py` applies it)

One JSON file per sheet keeps each step small: `A201.json`, `A202.json` … and `A100.json` for the RCM. A file holds only what changes. Rows are addressed by the numbers `show` prints for the **original** workbook; the script renumbers after additions and removals, so never adjust numbers yourself.

```json
{
  "sources": {
    "M1": "Walkthrough minutes, security deposits, Finance Manager and Leasing Manager, 17 Nov 2026 (minutes-deposits.txt)",
    "N":  "Staff notes in the workbook",
    "S1": "SOP Leasing v3.2, section 6 (SOP-Leasing.docx)"
  },
  "changes": [ { "sheet": "A203", "op": "…", "…": "…", "source": ["M1"], "basis": "Finance Manager, lines 44-50" } ],
  "open_points": ["SOP 6.4 says 90 days from the lease end; the Finance Manager said 30 days from handover. Confirm before Test 6."]
}
```
- `sources`: the same id for the same document in every file. Describe the document so a reviewer can find it (what, who was interviewed by role, date, file name).
- Every change needs `source` (one or more ids) and `basis`: where in the source the support is — the role who said it and the line or timestamp, or the document section; a short quote is best. No support, no change.
- `open_points`: contradictions between sources, things the walkthrough did not cover, documents still to obtain. They are printed and written to the log; they do not change the workbook. In `open_points` and `basis`, name a row by its control number or by the question's subject, not by its row number: row numbers move when rows are added or removed.
- A whole sheet the sources do not cover gets no update file: say so in one open point (`A201, A202 not covered by the minutes received`).

## Change types

| `op` | Address | Fields in `set` | Notes |
|---|---|---|---|
| `background` | `sheet` | — use `"text"` instead of `set` | `"mode": "append"` adds a paragraph (usual: what the walkthrough confirmed, changes in the period); default `replace` rewrites the whole background (give the full text) |
| `activity` | `sheet`, `"no"` | `activity`, `performed_by`, `frequency`, `control`, `owner`, `classification`, `type`, `de_finding`, `document` | give only the fields that change; a field you give replaces the whole cell |
| `add_activity` | `sheet`, `"after": <activity no>` or `"section": "<start of an existing section title>"` | all of the above; `de_finding` optional | the script assigns the next free control number (`A203-6`); you may ask for a number with `"control_ref"` in `set` if it is free. A new sub-section: `"section": "3.4 Title", "new_section": true` |
| `remove_activity` | `sheet`, `"no"` | — | only when the activity does not exist at this client. Its control number is taken out of the RCM and the questions automatically (logged) |
| `question` | `sheet`, `"no"` | `response`, and if needed `question`, `why`, `documents`, `interviewee`, `control_ref` | a response the staff already wrote is protected: `"mode": "append"` adds below it, `"mode": "replace"` rewrites it (only when the user asks) |
| `add_question` | `sheet`, optional `"after": <question no>` | `question`, `why`, `documents`, `interviewee` required; `response`, `control_ref` | for a point the team asked about that is not on the sheet |
| `risk` | `"sheet": "A100"` (or `"rcm"`), `"sn"` | `sub_process`, `risk`, `control_no`, `de_finding`, `audit_strategy`, `wp_ref` | `control_no` and `wp_ref` list every reference, one per line (`"A203-1\nA203-6"`) |
| `add_risk` | `"sheet": "A100"`, `"after": <risk sn>` or `"section": "<start of a section title>"` | all six fields | for a risk the walkthrough revealed that no row covers |
| `test` | `"sheet": "A308"` | `objective`, `basis`, `population`, `sample_size`, `source`, `steps`, `detail_title`, `detail_columns` (list), `n_rows` | only before any work is recorded on that sheet. `steps` is the full text (`Test 8: title\nCheck that:\n1. …`); the step columns and the RCM row quoting the full steps follow |

Control numbers of existing controls never change. Testing Result in the RCM is a link to the test's conclusion and is not set here.

- **New control numbers.** New controls take the next free numbers on their sheet in row order. When the RCM file must cite a new control, pin its number with `"control_ref"` in the `add_activity` (the next free number is one above the highest that `show` prints); `check` prints every number it assigns.
- **A control may sit under more than one risk**; list it under every risk it addresses. The order within `control_no` is free.
- **A question without a control** (`control_ref` left out) is fine when it concerns no control on the sheet.
- **Splitting a row.** Keep a library row whole when one person performs both parts and the outcome is the same. Split it when the parts have different owners or different design conclusions: the original row keeps its control number and the part that control describes; the other part becomes an `add_activity` with a new number.

## Wording
- **Activity**: who does what, when, in which system, with which document; present tense; one activity per row; roles, never names. Write what the walkthrough established, not what the library assumed.
- **Control**: what prevents or detects the risk and how it is evidenced. Where there is none: `No control is in place: <what happens instead>.`
- **Performed by / Owner**: role titles. **Frequency**: `As and when required`, `Daily`, `Weekly`, `Monthly`, `Quarterly`, `Annually`, `One-off`. **Classification**: `Preventive` / `Detective`. **Type**: `Manual` / `IT` / `Manual and IT`.
- **Walkthrough document**: change it only when a source shows that a document was sighted, obtained or promised (`Refund deposit forms for March 2027, sighted` · `AR ageing for September, requested from the Accounts Executive`); otherwise leave the planned document as it is.
- **Background**: `append` for what the walkthrough adds; `replace` (full text) when the existing background says something the sources contradict.
- **Design-effectiveness finding, per control**: `No design gap noted.` · `Design gap: there is no register of the expiry of banker's guarantees, so a guarantee may lapse before the lease ends, leaving the lease unsecured.` · `To be confirmed: the approval limits, pending the Delegation of Authority from the Head of Finance.` State the consequence; say when the practice differs from the SOP or the SOP is silent. One cell may combine them when a control is sound in one part and weak in another (`Design gap: … No design gap noted for the approval of the form. To be confirmed: …`).
- **Design-effectiveness conclusion, per risk (RCM)**: one conclusion over all the risk's controls — `No design gap noted.` or `Design gap: <gap> (A203-3). No other design gap noted.` Keep a prior-year observation that is already in the cell and add one sentence on whether the walkthrough shows it remediated (the script warns if you drop it).
- **Risk**: add the structural cause the walkthrough showed (small team, shared login, manual spreadsheet) to the existing sentence; keep "X may … This may result in …".
- **Response** added from the minutes: the substance in one to three sentences, starting `Per the minutes (<role>): …` when it is appended to a staff note. Not the transcript itself. A staff note that sits under the wrong question stays where it is: use it as evidence for the question it answers, and raise an open point so the team can move it.
- **Tests**: change a test only when the sources settle the fact (a step names a document or timeline the client does not have, or a gap just found has no step). When it depends on an open point, raise the open point and leave the test.
- No `[placeholders]` and no `{TOKENS}`: write the engagement's actual words. No client fact that is not in a source.

## Worked example — one sheet
The staff noted on question 2 that banker's guarantees are kept in a locked cabinet with no register; the minutes confirm it and add that two tenants paid a reduced deposit approved by the CEO by email; the SOP is silent on reduced deposits; nobody has ever forfeited a deposit.

```json
{
  "sources": {"M1": "Walkthrough minutes, security deposits, Finance Manager and Leasing Manager, 17 Nov 2026 (minutes-deposits.txt)", "N": "Staff notes in the workbook", "S1": "SOP Leasing v3.2, section 6 (SOP-Leasing.docx)"},
  "open_points": ["SOP 6.4 says refunds are made within 90 days of the lease end; the Finance Manager said 30 days from handover. Confirm which applies before Test 6."],
  "changes": [
    {"sheet": "A203", "op": "background", "mode": "append", "text": "The walkthrough on 17 November 2026 confirmed that about 70% of deposits are held in cash and 30% as banker's guarantees. Two tenants were given a reduced deposit of two months' rent in the period, approved by the CEO by email.", "source": ["M1", "N"], "basis": "Finance Manager, lines 40-58; staff note on question 1."},
    {"sheet": "A203", "op": "activity", "no": 1, "set": {"de_finding": "No design gap noted."}, "source": ["M1", "S1"], "basis": "Practice described by the Leasing Manager (lines 12-20) agrees with SOP 6.1."},
    {"sheet": "A203", "op": "activity", "no": 3, "set": {
        "activity": "Leasing keeps the original banker's guarantees in a locked cabinet. There is no register of expiry dates; the Leasing Manager checks the expiry of each guarantee ad hoc. Finance reconciles the security deposit balance in the property management system to the general ledger every month.",
        "control": "The security deposit listing is reconciled to the general ledger every month. The expiry of banker's guarantees is not tracked in a register.",
        "de_finding": "Design gap: there is no register or system alert for the expiry of banker's guarantees, so a guarantee may lapse before the lease ends, leaving the lease unsecured."},
      "source": ["M1", "N"], "basis": "Leasing Manager, line 72: 'we don't have a list, I look at them when I remember'; staff note on question 2."},
    {"sheet": "A203", "op": "add_activity", "after": 2, "set": {
        "activity": "Where a deposit below three months' rent is proposed, the Head of Leasing submits the request to the CEO, who approves it by email before the Tenancy Agreement is signed.",
        "performed_by": "Head of Leasing and CEO", "frequency": "As and when required",
        "control": "A reduced security deposit is approved by the CEO in writing before the Tenancy Agreement is signed.",
        "owner": "CEO", "classification": "Preventive", "type": "Manual",
        "de_finding": "Design gap: the Delegation of Authority names no approver for reduced deposits; the CEO's email approval is a practice, not a documented requirement.",
        "document": "CEO approval emails for the two reduced deposits (to be obtained)"},
      "source": ["M1", "N"], "basis": "Finance Manager, lines 44-50; SOP section 6 is silent."},
    {"sheet": "A203", "op": "remove_activity", "no": 5, "source": ["M1"], "basis": "Finance Manager, line 90: no deposit has ever been forfeited and there is no forfeiture process; arrears are deducted on the refund form (activity 4)."},
    {"sheet": "A203", "op": "question", "no": 1, "mode": "append", "set": {"response": "Per the minutes (Finance Manager): the two reduced deposits were approved by the CEO by email; the Delegation of Authority names no approver for this."}, "source": ["M1"], "basis": "Lines 44-50."},
    {"sheet": "A203", "op": "question", "no": 3, "set": {"response": "Finance reconciles the deposit listing to the general ledger monthly; the Finance Manager signs the reconciliation."}, "source": ["M1"], "basis": "Lines 61-66."}
  ]
}
```
and in `A100.json`, for the risks `show` listed under that sheet:
```json
{"sheet": "A100", "op": "risk", "sn": 11, "set": {"control_no": "A203-1\nA203-2\nA203-6", "de_finding": "Design gap: reduced deposits are approved by the CEO by email (A203-6) but the Delegation of Authority names no approver for them. No other design gap noted."}, "source": ["M1", "N"], "basis": "Finance Manager, lines 40-58."}
```

## What `check` and `apply` refuse (correct the file; do not work around)
- a row number that does not exist, or appears twice on the sheet; a field that is not in the table above; the same cell changed twice;
- a change without `source` or `basis`, or with a source id the file does not declare;
- overwriting a staff response without `mode`;
- a new activity or question with required fields missing (they would show as `[placeholders]`);
- a control number or test reference that does not exist after the update; a control number used twice;
- text that still holds a `{TOKEN}` or a `[placeholder]`;
- a change to a test sheet on which work is already recorded;
- writing over the workbook you were given.

Warnings to read: a design-effectiveness cell that loses its prior-year observation; a risk left without a control; an RCM audit strategy that quotes steps differently from its test; classification or type outside the house values.

## What the script guarantees
- Only the sheets with changes are rewritten, with the same builders as awp-generate: same columns, merges, fonts, borders, widths and row heights. Scope, the index sheets, untouched walkthrough sheets, test sheets and the DA sheet are not touched.
- Everything in the tables of a rewritten sheet that you did not change — including every staff response — is carried over unchanged. Anything outside the tables on such a sheet is listed and the update stops unless the user accepts the loss.
- The hidden `_build` sheet gets a line recording the update. `verify` (the same checks as awp-generate) runs on the result.
- The update log has five sheets: Changes (sheet, item, field, before, after, source, basis), Sources, Open points, Still open, Warnings.
