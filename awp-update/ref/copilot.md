# awp-update — setting it up and running it in corporate Copilot (Microsoft 365 Copilot Agent Builder)

## Set-up (once)
1. Open the AWP agent in Agent Builder → Configure → Skills → Add → Upload skill → `copilot-awp-update-<version>.zip` (upload the zip as it is). It sits beside `awp-generate` in the same agent; an agent takes up to eight skills.
2. Paste the whole of `ref/agent-instructions.txt` (in this zip; in the repository, `awp-agent/instructions.txt`) into the agent's Instructions, replacing what is there. The same text serves every `awp-*` skill.
3. In Preview: "run the awp-update selftest" → `SELFTEST PASSED - awp-update v2.4 · engine awp-generate v2.3`.

## Per engagement (what the user does)
1. Attach to the chat: the current AWP workbook (`.xlsx`, with the staff's notes typed into the walkthrough sheets), the walkthrough minutes (`.txt`), and any documents received (SOP, policies, Delegation of Authority).
2. Prompt, for example:
   > The walkthroughs are done. Update the 100 and 200 series of the attached AWP from the walkthrough minutes, the staff's notes in the workbook and the SOP. Show me the design gaps and open points before you apply.
3. Review the gaps and open points, say "apply", and download the two files from the links in the reply: `<name>_updated.xlsx` and `<name>_updated_log.xlsx`. If the reply shows a path instead of a link, say "give me both files as download links". Review the log before relying on the workbook.

## Keeping a long update from timing out
- The skill already works one walkthrough sheet per step. For a large AWP, drive it in turns: "Do A201 and A202 now", then "continue with A203 and A204", then "now the RCM and apply".
- If the sandbox does not keep files between turns, ask the model to print each finished update file (`A201.json` …) in the chat as it goes, and to re-create them from the chat before `apply`. Re-attach the workbook if the model says it is gone.
- Many long transcripts: attach the minutes for the sheets being done in that turn.
- The format never depends on the model; a stronger model ("Think deeper") gives better judgement on gaps.

## Limits to know
- Close the workbook in Excel and attach a saved copy. `.xlsx` only.
- The 100 and 200 sheets are edited in place, so the team's notes, comments and highlighting stay. What no script can keep when it saves a workbook: embedded files (documents inserted as objects) and Excel 365 threaded comments, anywhere in the workbook. The skill lists them and stops so that the user can decide — detach them first or keep them in the original.
- Test sheets are changed only when the walkthrough evidence requires it and only before work is recorded on them; such a sheet is laid out again, so anything outside its tables is listed first.
