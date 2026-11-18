# etcmail — changelog

| Date | Version | Change |
|---|---|---|
| 2026-11-18 | 1.4 | One full-page screenshot per engagement is enough: the script cuts it into the house crops (header down to just above "Expand to view metrics"; a full-width strip around the red-bordered Change in Revenue Recognition box) with Pillow/numpy, falling back to the full screenshot with a warning. |
| 2026-11-18 | 1.3 | No signature: the draft ends at "Best regards," and users add their own Outlook signature; `sender` dropped from the settings block. Every run of text is Arial 11 pt explicitly. |
| 2026-11-18 | 1.2 | Screenshots that never reach the sandbox are caught instead of silently dropped: new `files` command (lists the image files the sandbox really has, with shape and fingerprint), duplicate pictures and a missing header screenshot are warned, the reply reports `Screenshots embedded: X of M` with the fix. One screenshot per engagement (header down to the revenue-recognition box) is accepted. |
| 2026-11-18 | 1.1 | Reusable for anyone: no personal details in the skill. Partner, CC and signature come from an `etcmail settings` block in the agent's Instructions (copied into the input as `"settings"`), or `--settings FILE`. Removed `lib/` and the `settings` command; no signature logo. Copilot zip committed to `dist/`. |
| 2026-11-18 | 1.0 | First release: screenshots → input JSON → Outlook draft (.eml) in the house format (five-column table, red/green revenue change, screenshots embedded, signature); Copilot Agent Builder setup in `ref/agent-builder.md`. Checked by rebuilding two real approval emails. |
