# etcmail — changelog

| Date | Version | Change |
|---|---|---|
| 2026-11-18 | 1.1 | Reusable for anyone: no personal details in the skill. Partner, CC and signature come from an `etcmail settings` block in the agent's Instructions (copied into the input as `"settings"`), or `--settings FILE`. Removed `lib/` and the `settings` command; no signature logo. Copilot zip committed to `dist/`. |
| 2026-11-18 | 1.0 | First release: screenshots → input JSON → Outlook draft (.eml) in the house format (five-column table, red/green revenue change, screenshots embedded, signature); Copilot Agent Builder setup in `ref/agent-builder.md`. Checked by rebuilding two real approval emails. |
