# awp-update — setting it up and running it in corporate Copilot (Microsoft 365 Copilot Agent Builder)

## Set-up (once)
1. Open the AWP agent in Agent Builder → Configure → Skills → Add → Upload skill → `copilot-awp-update-<version>.zip` (upload the zip as it is). It sits beside `awp-generate` in the same agent; an agent takes up to eight skills.
2. Replace the agent's Instructions with:

```
You help internal auditors with audit work programs (AWP) using two skills.
- awp-generate: to prepare or build an AWP, work program, workpapers, RCM, walkthrough or test sheets, or the blank template.
- awp-update: after the walkthroughs, to update the RCM (100) and the walkthrough sheets (200 series) of an existing AWP from walkthrough minutes, the staff's notes in the workbook and supporting documents such as SOPs.
Read the SKILL.md of the skill that fits and follow it exactly. Run that skill's selftest once per conversation before the first real command and report the versions. Workbooks are produced and changed only by the skills' scripts; never write, restyle or hand-patch a workbook yourself. For an update, work one sheet at a time and never paste a whole sheet into the chat. If a selftest fails or a script cannot run, say so and stop rather than improvising. Deliver the workbook with the script's verification result and the open points.
```
3. In Preview: "run the awp-update selftest" → `SELFTEST PASSED - awp-update v1.0 · engine awp-generate v2.1`.

## Per engagement (what the user does)
1. Attach to the chat: the current AWP workbook (`.xlsx`, with the staff's notes typed into the walkthrough sheets), the walkthrough minutes (`.txt`), and any documents received (SOP, policies, Delegation of Authority).
2. Prompt, for example:
   > The walkthroughs are done. Update the 100 and 200 series of the attached AWP from the walkthrough minutes, the staff's notes in the workbook and the SOP. Show me the design gaps and open points before you apply.
3. Review the gaps and open points, say "apply", and download the two files: `<name> (updated).xlsx` and `<name> (updated) - update log.xlsx`. Review the log before relying on the workbook.

## Keeping a long update from timing out
- The skill already works one walkthrough sheet per step. For a large AWP, drive it in turns: "Do A201 and A202 now", then "continue with A203 and A204", then "now the RCM and apply".
- If the sandbox does not keep files between turns, ask the model to print each finished update file (`A201.json` …) in the chat as it goes, and to re-create them from the chat before `apply`. Re-attach the workbook if the model says it is gone.
- Many long transcripts: attach the minutes for the sheets being done in that turn.
- The format never depends on the model; a stronger model ("Think deeper") gives better judgement on gaps.

## Limits to know
- Close the workbook in Excel and attach a saved copy. `.xlsx` only.
- Embedded files, pictures, cell comments and notes typed outside the tables on a sheet that changes cannot be carried over by a script; the skill lists them and stops so that the user can decide. Sheets that do not change are not affected, except that embedded files are dropped from the whole workbook by any script that saves it — detach them first or keep them in the original.
- Test sheets are changed only on request of the walkthrough evidence and only before work is recorded on them.
