# awp-generate — sheet layout and formatting (enforced by `scripts/awp.py`; do not restyle by hand)

## Global
- Arial 8 everywhere. Gridlines off. Tab colour 92D050 (Scope tab black).
- Row 1, last used column: the sheet reference in a yellow (FFFF00) box, bold, centred, double border.
- A3:B5: `Client`, `Process`, `Audit period` (RCM uses `Client:` / `Process:` / `Review period:`). Values are typed on the index sheet only; every other sheet uses `='A'!B3` … so a change propagates.
- Index sheet also carries `Prepared by:` / `Date:` (row 3) and `Reviewed by:` / `Date:` (row 5) in D-G, and the build stamp in grey italic two rows under the last index item.
- Cross-references (control numbers, WP refs, index refs) are bold red (FF0000), centred.
- Column headers: grey fill D9D9D9, bold, centred (left for text-heavy columns), wrapped, thin borders.
- Text cells: wrapped, top-aligned, thin borders; row height sized to the longest text in the row (11.25 pt per line).
- Hidden sheet `_build`: skill version, build date, library versions, module ids, tokens, stamp text. Six short rows; no effect on size or speed.

## Scope
Columns `No | Key Process | Indicative area of focus | Covered?`. Header fill black with white bold text; `Covered?` header fill 4BACC6. One row per area; `No` and `Key Process` merged down the key process. Widths 6 / 49 / 80 / 42.

## Index sheets (A, A200, A300)
Row 7: title (`Index`, `Process walkthrough`, `Testing`) bold with a bottom border in B, `Ref` in the ref column. Items every second row from row 9: number centred in A, title in B (width 64), ref bold red in the ref column (G for A and A200, E for A300).

## A100 — Risk and control matrix
Header row 7: `S/N | Sub-process | Risk | Control No. | Finding for Design Effectiveness | Audit Strategy | Testing Result | WP Ref` (height 30.75). Section rows merged A:H, fill DCE6F1, bold (`1. Overall control environment`). Data rows: S/N centred; sub-process, risk, DE finding, audit strategy, testing result left-wrapped; control no. and WP ref bold red centred. `Testing Result` holds a formula to the Conclusion cell of the first WP-ref test sheet (blank until concluded) unless text is supplied. Freeze panes C8. Widths 14.4 / 44.9 / 69.6 / 15.9 / 67.3 / 83.1 / 70.9 / 13.6.

## A201… — Walkthrough
- Row 7 merged A:K: `Subprocess title: <title>` — fill 808080, white bold.
- Row 8 merged A:K: `Background:` + blank line + narrative (process, systems, people; changes in the period; prior-year context).
- Row 9 header: `No. | Activity | Performed By | Frequency | Control Reference No. | Existing controls / mitigating measures | Control Owner | Control Classification | Control Type | Finding for design effectiveness | Walkthrough Document`.
- Sub-section rows merged A:K, bold, no fill (`1.1 Policies and procedures and Delegation of Authority`). Activity rows numbered continuously across sub-sections; control refs `<sheet>-<n>` bold red; a row may simply cross-refer (`Please refer to <A302> …`).
- Blank row, then merged A:K band (808080, white bold) `Walkthrough questions for the team lead (…)`, header `No. | Question (B:D) | Why we are asking, including the prior-year context (E:F) | Documents to request during the session (G:H) | Proposed interviewee | Response noted during the walkthrough | Control Reference No.`; question rows keep the same merges.
- Freeze panes A10. Widths 12.6 / 42.6 / 19 / 19.1 / 16 / 50.4 / 16.1 / 16.6 / 12.9 / 59.1 / 16.1.

## A301… — Test sheet
- Row 7: `Test n` in A and the title in B on an 808080 band with white bold text across the table width.
- Rows 9-19 (every second row): `Objective:`, `Basis of Selection:`, `Population:`, `Sample Size:`, `Source:`, `Audit Steps:` — label bold in A, text merged B:I.
- Row 21: `Work done:` in A; detail-title header merged over the sample-detail columns (fill B7DEE8); `Work Done` merged over the step columns (fill 92D050); `Remarks` over the last column (B7DEE8).
- Row 22: `S/N`, the detail column names (B7DEE8), step numbers 1..n (92D050), `Exceptions and remarks` (B7DEE8); height 45.
- Rows 23+: S/N 1..n with bordered empty cells (height 33.75). Computed columns carry a formula per row (days between dates, rent × n, variance %).
- Then `Notes:` / `Findings:` / `Conclusion:` every second row (defaults `Nil`, `Nil`, `To be concluded on completion of fieldwork.`). The Conclusion cell is what the RCM links to. During fieldwork: notes numbered N1, N2 … referenced from the remarks column; conclusion `No exceptions noted.` / `Exceptions noted — refer to Findings.`
- Freeze panes C23. Widths A 20.6, B 7.6, detail columns 14-42 by header length, step columns 8.9, remarks 47.4.

## A400 — Data analytics (legacy name `DA` is accepted on export)
Row 7 `Data analytics` / `Ref`. Header row 9: `S/N | Analytic | Data source | What the analytic is designed to identify | Feeds | Result | Status`. `Feeds` bold centred (test refs, newline-separated when several). Note row merged A:G below the table. Widths 13.6 / 64 / 18.9 / 60 / 9.7 / 30 / 11.
