# etcmail — maintainers' notes

Current: **etcmail v1.0**. `VERSION` lives in `scripts/etcmail.py`; history in `CHANGELOG.md` and GitHub.

## Layout
```
etcmail/
  SKILL.md                    instructions for the model (Copilot agent or Claude) - the entry point
  README.md, CHANGELOG.md     these notes
  scripts/etcmail.py          the engine: input JSON -> Outlook draft (.eml, X-Unsent) + summary; standard library only
  lib/settings.example.json   placeholder settings (committed)
  lib/settings.json           the owner's real settings: sender signature, partners and their emails, CC (NOT committed)
  lib/logo.gif                signature logo (NOT committed)
  ref/agent-builder.md        the Copilot Agent Builder fields (name, description, instructions, starter) and daily use
```
The repository is public, so `lib/settings.json` and `lib/logo.*` are git-ignored. The script uses `settings.json` when present, otherwise the example (and says so in every build).

## Changing recipients or the signature
Edit `lib/settings.json` (`default_partner`, `partners` name → email, `cc`, `sender`, optional `client_short` map for subject names), run `python scripts/etcmail.py selftest` and `settings`, then release.

## Release
1. `python etcmail/scripts/etcmail.py selftest` → `SELFTEST PASSED`, naming `settings.json` (not the placeholder).
2. `python release.py etcmail <version>` from the repository root → `dist/copilot-etcmail-<version>.zip` (flat, `SKILL.md` at the root — the one to upload in Copilot Agent Builder) and `dist/skill-etcmail-<version>.zip`.
3. These zips contain the real settings: keep them out of git (`dist/` is ignored; do not force-add them) and hand them to the user directly.
4. Re-upload the zip in Agent Builder (Skills → replace) and test with one pair of screenshots.

## Design decisions
- The model reads the screenshots (judgement); the script owns every format rule (subject, greeting, table, number format and colour, image embedding, signature), so the email is identical whatever model Copilot uses.
- `.eml` with `X-Unsent: 1` rather than `.msg`: plain standard library, and classic Outlook opens it as an editable draft with the images inline.
- Screenshots are ordered within an engagement by shape (the flat revenue-recognition box goes second), so upload order inside a pair does not matter.
- Partner is not on the screenshots: it comes from `default_partner`, overridable per engagement from the user's words.
