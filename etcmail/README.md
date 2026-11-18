# etcmail — maintainers' notes

Current: **etcmail v1.5**. `VERSION` lives in `scripts/etcmail.py`; history in `CHANGELOG.md` and GitHub.

## Layout
```
etcmail/
  SKILL.md                    instructions for the model (Copilot agent or Claude) - the entry point
  README.md, CHANGELOG.md     these notes
  scripts/etcmail.py          the engine: input JSON -> Outlook draft (.eml, X-Unsent) + summary; standard library only
  ref/agent-builder.md        Agent Builder fields (name, description, instructions with the settings block, starter) and daily use
```
The skill holds no personal details, so the same `dist/copilot-etcmail-<version>.zip` serves everyone. Each user's partner
and CC live in the `etcmail settings` JSON block of their own agent's Instructions; the model copies it into the
input as `"settings"`. Outside Copilot, `build ... --settings FILE` reads the same block from a file (keep such files out of git).

## Release
1. `python etcmail/scripts/etcmail.py selftest` → `SELFTEST PASSED`.
2. `python release.py etcmail <version>` from the repository root → `dist/copilot-etcmail-<version>.zip` (flat, `SKILL.md` at the root — the one to upload in Copilot Agent Builder) and `dist/skill-etcmail-<version>.zip`.
3. Commit the Copilot zip (`git add -f dist/copilot-etcmail-<version>.zip`; remove the previous version); users re-upload it in Agent Builder (Skills → replace). Their settings stay in their instructions.

## Design decisions
- The model reads the screenshots (judgement); the script owns every format rule (subject, greeting, table, number format and colour, image embedding), so the email is identical whatever model Copilot uses.
- `.eml` with `X-Unsent: 1` rather than `.msg`: plain standard library, and classic Outlook opens it as an editable draft with the images inline.
- A full-page screenshot is cut by the script, not the model: the light-blue revenue-recognition box and the text bands of the four tiles are found by pixel, so the crops match the house emails every time.
- Screenshots are ordered within an engagement by shape (the flat revenue-recognition box goes second), so upload order inside a pair does not matter.
- Partner and CC are not on the screenshots and differ per user: they come from the settings block in the agent's Instructions, not the skill, so the zip is reusable and holds nothing personal. The partner can be overridden per engagement from the user's words.
- No signature: the draft ends at "Best regards," and each user adds their own Outlook signature, so the agent works for anyone without signature settings.
