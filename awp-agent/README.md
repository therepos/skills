# awp-agent — the AWP agent for Microsoft 365 Copilot (Agent Builder)

Not a skill: the settings that turn the `awp-*` skills into one agent. Everything needed to set the agent up in any Microsoft 365 tenant that has Copilot Agent Builder with skills is in this repository — nothing is kept only in someone's agent.

| Setting | Value |
|---|---|
| Name | `AWP` |
| Description | `Build and update internal-audit work programs` |
| Instructions | the whole of [`instructions.txt`](instructions.txt), pasted as it is |
| Skills | `dist/copilot-awp-generate-<version>.zip` and `dist/copilot-awp-update-<version>.zip` (the newest of each) |
| Knowledge | none. Leave the knowledge sources as they are; the skills read only what the user attaches |

## Set up in a new environment
1. Download the two `copilot-awp-*.zip` files from `dist/` and open `instructions.txt`.
2. Copilot → Agents → Create agent → Configure. Enter the name and description.
3. Paste `instructions.txt` into **Instructions**.
4. **Skills → Add → Upload skill**: upload each zip as it is (do not unzip). The list then shows each skill with its version at the start of the description (`v2.3 - Build …`).
5. Save, then in **Preview** ask: `run the selftest of each skill`. Expect two lines starting `SELFTEST PASSED`, with the same versions as the zips.
6. If a selftest cannot run, the chat has no script sandbox: try another model in the model picker, in a new chat. The skills need Python with `openpyxl`, and the sandbox cannot install packages.

## Updating
- A new skill release: delete that skill in the agent, upload the new zip. A skill without a version in front of its description is an old upload.
- A change to the instructions: edit `instructions.txt` here, then paste it again in every environment. `release.py` copies it into each `awp-*` zip as `ref/agent-instructions.txt`, so a zip always carries the instructions it was released with.

## What the instructions do, line by line
- Name both skills and when each applies, so a request to update is not answered by building.
- Selftest first: proves the sandbox can run the scripts and states the versions in the reply.
- Scripts only: the format lives in the scripts; a model that writes cells itself drifts from it.
- One sheet at a time: long updates time out otherwise.
- Download links: a path printed as text is not something the user can open.
