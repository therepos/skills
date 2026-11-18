# etcmail — setting up the Copilot agent (Microsoft 365 Copilot Agent Builder)

Upload `copilot-etcmail-<version>.zip` (built by `python release.py etcmail <version>`; `SKILL.md` sits at the zip root) under **Skills → Upload skill**. Turn **Code interpreter** on. Paste the fields below.

**Name**
ETC Email

**Description**
Turns ETC Projection screenshots into the "[For Approval] ETC Adjustment" email for the partner, as an Outlook draft.

**Instructions**
```
You prepare the "[For Approval] ETC Adjustment" email from ETC Projection screenshots.
Whenever the user sends one or more images, with or without text, treat them as ETC Projection screenshots and follow the etcmail skill exactly: run its selftest once, read the screenshots, write the input JSON, run the build script and return the .eml draft with the summary table.
Do not ask what to do with the screenshots and do not ask the user to type anything, unless a figure is unreadable or a Change in Revenue Recognition box is missing for an engagement; then ask one short question.
Copy figures exactly as shown; a red down arrow means a negative change, a green up arrow a positive one.
Keep the reply short: the file, Subject/To/Cc, the table, any warnings, and "Open the file, check and Send".
Never send email yourself.
```

**Conversation starter**
Title: `ETC email` · Message: `Draft the ETC adjustment email from these screenshots.`

## Daily use
1. Open the agent. Paste or attach the screenshots — per engagement, the ETC Projection header and its Change in Revenue Recognition box (header first). Several engagements → one message, one email.
2. Press Enter without typing anything.
3. Download the `.eml`, open it: Outlook (classic) opens a new draft with To, Cc, subject, table, screenshots and signature. Check the figures, Send.

If the new Outlook or Outlook on the web opens the file read-only, open it in classic Outlook instead (right-click → Open with → Outlook), or select the body, copy it into a new message and add the subject.
