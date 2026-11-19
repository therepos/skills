---
name: awp-update
description: Update an internal-audit work program (AWP) workbook after the walkthrough - bring the walkthrough sheets (200 series) and the risk and control matrix (100) in line with what the client actually does, from the walkthrough minutes, the staff's notes in the workbook and supporting documents such as SOPs and policies. Use when an auditor says the walkthroughs or interviews are done and asks to update the AWP, the RCM, the walkthrough sheets, the process documentation or the design-effectiveness findings, or attaches walkthrough minutes or a transcript together with an AWP workbook.
---

# awp-update — update the RCM and the walkthrough sheets after the walkthrough

**Use this skill when** the walkthrough interviews are done and the auditor wants the AWP brought up to date: the 200-series walkthrough sheets (background, activities, controls, design-effectiveness findings, responses) and the 100 risk and control matrix (risks, control numbers, design-effectiveness conclusions). Building a new AWP is the **awp-generate** skill; harvesting a finished AWP into the libraries is **awp-learn**.

**What it produces:** a new copy of the workbook, `<name> (updated).xlsx`, in the same house format, plus `<name> (updated) - update log.xlsx` listing every changed cell with its text before and after, the source and the basis. Only the sheets that change are rewritten; every other sheet stays exactly as the team left it. The workbook the user gave you is never overwritten.

**Division of work.** You read the sources and decide what changes: that is judgement. `scripts/update.py` does everything else: it writes the cells, keeps the format, numbers new controls, renumbers rows, keeps the RCM and the question references consistent, protects the staff's notes and writes the log. Never open the workbook with openpyxl to write cells yourself, never restyle it, and never paste a whole sheet into the chat: that is what times out.

---

## The user says… → what you do
| The user says something like | Do |
|---|---|
| "the walkthroughs are done, update the AWP / the RCM / the 200 series", "update the workpaper from these minutes", attaches minutes or a transcript with an AWP | the `/awp-update` procedure below |
| "also use the SOP / policy / these documents we received" | same procedure; the documents are further sources |
| "prepare / build the AWP", "give me the template" | not this skill: **awp-generate** |
| "save what we learned for the next team" | not this skill: **awp-learn** |
| "change the format / add a column / fix the script" | decline: the format belongs to the skill owners |

## The sources and what each one proves
Give every source a short id and use the same id in every update file.
- **M1, M2 … walkthrough minutes** (verbatim transcripts, `.txt`): what the client says it actually does. This is the evidence for how the process and the controls operate. Note who said it (by role) and where (line or timestamp).
- **N — the staff's notes in the workbook**: the "Response noted during the walkthrough" column of each 200 sheet. They are the audit team's own record. Read them through `show`; treat them as evidence equal to the minutes.
- **S1, S2 … supporting documents** (SOP, policy, Delegation of Authority, organisation chart, system list): what is *supposed* to happen. They prove the designed process, not the practice.

Where the sources disagree, do not choose silently. Describe the actual practice in the activity and control, state the difference in the design-effectiveness finding (practice differs from the SOP, or the SOP is silent), and list it under `open_points` when the team must confirm it. Where the sources say nothing about a row, leave the row unchanged and list it as not covered: never conclude without evidence, never invent a client fact.

## `/awp-update` — the procedure
Work inside the folder that contains `SKILL.md`. Keep the engagement's files outside it: `../work/` for the workbook, the sources and the update files, `../out/` for what you deliver. Run `python scripts/update.py selftest` once per session (`python3` if `python` is not found); it must print `SELFTEST PASSED - awp-update vN · engine awp-generate vN`. Report the versions. Needs Python 3 and `openpyxl`, nothing else.

1. **Collect.** The current AWP workbook (the one the staff typed their notes into), the walkthrough minutes, any supporting documents. If the workbook is missing, ask for it: nothing can be updated without it. Assign the source ids.
2. **Overview.** `python scripts/update.py show "../work/<AWP>.xlsx"` prints the sheets, the counts and what is still open (questions without a response, controls and risks without a design conclusion).
3. **Read the sources in full** before writing anything. For each walkthrough sheet, note which parts of the minutes and which sections of the documents concern it.
4. **One walkthrough sheet at a time** (this is what keeps a long job from timing out; finish and check one sheet before starting the next):
   a. `python scripts/update.py show "../work/<AWP>.xlsx" A202` prints that sheet: background, every activity row with its control and number, every question with the staff's response, and the RCM risks that rely on its controls.
   A sheet the sources do not cover is left alone: no update file, one open point naming it.
   b. Compare each row with the sources and write `../work/A202.json` — only the changes, in the format in `ref/update.md` (read it before the first file; it has the fields, the wording rules and a worked example).
   c. `python scripts/update.py check "../work/<AWP>.xlsx" ../work/A202.json` → correct every `error:`; read every `warn:`. It must end `CHECK PASSED`.
5. **Then the RCM.** Each sheet's `show` already prints the risks that rely on its controls; `python scripts/update.py show "../work/<AWP>.xlsx" rcm` prints all of them. Write `../work/A100.json`: for each risk whose controls you assessed, the design-effectiveness conclusion, the control numbers when controls were added, and the risk wording when the walkthrough showed a cause the risk does not name. Run `check` with **all** the files together (the RCM refers to controls added by the sheet files): `check "<AWP>.xlsx" ../work/A201.json ../work/A202.json … ../work/A100.json`.
6. **Apply everything in one go:** `python scripts/update.py apply "../work/<AWP>.xlsx" "../out/<AWP> (updated).xlsx" ../work/A201.json … ../work/A100.json`. It must print `VERIFY PASSED` and `UPDATE APPLIED`. All files are applied against the row numbers of the original workbook, so never apply file by file onto an already updated copy.
7. **Deliver** from `../out/`: the updated workbook and the update log. Report in a few lines: sheets updated and cells changed; design gaps found (control number and one line each); where practice differs from the SOP; the open points; what is still open (the script prints it). Say that the team should review the update log before relying on the workbook.

If the user wants it in one go, run straight through. Otherwise stop after step 5 and show the gaps and open points before applying. If a session is cut off, the update files already written in `../work/` are the progress: continue with the next sheet.

### When the script says something would be lost
`apply` rewrites a changed sheet from its table content. If a changed sheet holds something outside the tables — a cell comment, a note typed beside the table, a picture — or the workbook has embedded files, the script lists each item and writes nothing. Tell the user exactly what was listed and let them choose: move it into a table cell in Excel and run again, or accept the loss (`--force`). Never use `--force` on your own decision. If the user cannot accept either, deliver the update files' content as a list of changes for the team to type in.

---

## Judgement rules
- **Staff responses are never rewritten.** A blank response may be filled from the minutes. Where the staff already wrote one, add to it only when the minutes give something the note lacks (`"mode": "append"`, starting `Per the minutes (<role>): …`). Use `"mode": "replace"` only when the user asks for it.
- **Expected becomes actual.** Before the walkthrough the sheets describe the process we assumed, often labelled as such (`Expected process`, `Expected: …`, `Expected control: …`). Every row, background and section title the sources cover is rewritten as the actual process and the label goes; `show` marks what still carries it (`PLANNING WORDING`) and the script refuses new text that keeps it. A row the sources do not cover keeps its planning wording — that is the honest state — and is reported as still open.
- **Activities and controls describe this client, as fact.** An activity says who does what, when, in which system and with which document. A control says who checks or approves what, against what, how often, and what evidence it leaves. Present tense; roles, never people's names (the minutes name people; write their role). Not "should", "is expected to" or "will": what ought to happen but does not belongs in the design-effectiveness finding. The full pattern and a before/after are in `ref/update.md`.
- **A planned control the client does not have:** keep the row, say in the control cell what is and is not in place, and record the design gap. Remove a row (`remove_activity`) only when the activity itself does not exist at this client.
- **A control the client has that the sheet lacks:** add it (`add_activity`); the script gives it the next free control number. Add it to the RCM risk it addresses, and add a question row if the team asked about it.
- **Design effectiveness, per control** (walkthrough sheet): `No design gap noted.` · `Design gap: <what is missing or weak> so that <consequence>.` · `To be confirmed: <what is outstanding and from whom>.`
- **Design effectiveness, per risk** (RCM): one conclusion over the risk's controls, naming the control numbers where there is a gap. Keep any prior-year observation that is in the cell and add whether the walkthrough shows it remediated.
- **Tests** (300 series) change only when the walkthrough shows that a planned step cannot be performed or misses a gap just found: use the `test` change; the RCM row that quotes the test's steps follows automatically. A test sheet with work already recorded cannot be changed here.
- Every change carries its `source` ids and a `basis` (who said it, or the document section). A change you cannot support is not made.

## Files
- `SKILL.md` — this file. `README.md` — maintainers' notes; `CHANGELOG.md`.
- `scripts/update.py` — selftest / show / check / apply. `scripts/awp.py` — the shared engine from awp-generate (the house format), bundled into the release zip.
- `ref/update.md` — the update file: every change type and field, wording rules, a worked example.
- `ref/copilot.md` — setting this up and running it in corporate Copilot.
