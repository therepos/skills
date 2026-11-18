# etcmail — setting up your Copilot agent (Microsoft 365 Copilot Agent Builder)

The skill zip is the same for everyone and holds no personal details. Your own details — partner, CC, signature — go in your agent's **Instructions**, which stay private to you in your organisation's tenant.

1. Download `dist/copilot-etcmail-<version>.zip` from the repository (`SKILL.md` sits at the zip root).
2. In Agent Builder: create an agent, **Skills → Upload skill** → the zip. Turn **Code interpreter** on.
3. Paste the fields below. In the Instructions, replace every `<…>` in the settings block with your own details.

**Name**
ETC Email

**Description**
Turns ETC Projection screenshots into the "[For Approval] ETC Adjustment" email for the partner, as an Outlook draft.

**Instructions**
```
You prepare the "[For Approval] ETC Adjustment" email from ETC Projection screenshots.
Whenever the user sends one or more images, with or without text, treat them as ETC Projection screenshots and follow the etcmail skill exactly: run its selftest once, read the screenshots, write the input JSON, run the build script and return the .eml draft with the summary table.
Always put the etcmail settings block below, unchanged, into the input JSON as "settings".
Do not ask what to do with the screenshots and do not ask the user to type anything, unless a figure is unreadable or a Change in Revenue Recognition box is missing for an engagement; then ask one short question.
Copy names, codes and figures exactly as shown; a red down arrow means a negative change, a green up arrow a positive one.
Keep the reply short: the file, Subject/To/Cc, the table, any warnings, and "Open the file, check and Send".
Never send email yourself.

etcmail settings:
{
  "sender": {
    "name": "<Your Name>",
    "title": "<Service Line>",
    "company": "<Company Pte. Ltd.>",
    "address": "<Office address>",
    "mobile": "<+65 0000 0000>",
    "email": "<your.email@company.com>",
    "website": "<www.company.com>"
  },
  "default_partner": "<Partner Name>",
  "partners": {
    "<Partner Name>": "<partner.email@company.com>"
  },
  "cc": ["<Manager Name> <manager.email@company.com>"],
  "client_short": {}
}
```

Settings fields:
- `default_partner` — the Partner In-Charge for every row, the To recipient and the "Hi <first name>," greeting. Must also appear in `partners` with an email.
- `partners` — name → email for every partner you may name in a message ("partner is Jane Tan"); add more lines as needed.
- `cc` — list of `Name <email>`; `[]` for none.
- `client_short` — optional subject names when the automatic ones are not what you use, e.g. `{"ACME CORPORATION LIMITED": "Acme"}`. Automatic: `MINISTRY OF HEALTH` → `MOH`, `ACME CORPORATION LIMITED` → `Acme`.
- `sender` — the text signature under "Best regards,". Leave a field out to omit that line.

**Conversation starter**
Title: `ETC email` · Message: `Draft the ETC adjustment email from these screenshots.`

## Daily use
1. Open the agent. **Attach** the screenshots as files (📎) — pasting several images can leave only the last one in Copilot's file sandbox. Per engagement: the ETC Projection header and its Change in Revenue Recognition box (header first), or one screenshot covering both. Several engagements → one message, one email.
2. Press Enter without typing anything. To name a different partner for one email, type it ("partner is Jane Tan; previously Lee Lead").
3. Download the `.eml` and open it: classic Outlook opens a new draft with To, Cc, subject, table, screenshots and signature. Check the table against the screenshots, then Send.

If the new Outlook or Outlook on the web opens the file read-only, open it in classic Outlook instead (right-click → Open with → Outlook), or select the body, copy it into a new message and add the subject.
