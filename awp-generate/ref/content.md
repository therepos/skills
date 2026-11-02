# awp-generate — how to write each element (generic house examples; tokens shown as they appear in the libraries)

The libraries already carry this depth. This guide is for the content you add in the spec during `/awp-build` and for the text you generalise during `/awp-learn`.

## Scope — indicative areas of focus
Each line is a sentence beginning "Check that …" (or "Assess whether …") that names the control objective and, where relevant, the system, document or period-specific event.
> Check that security deposits and banker's guarantees are collected in accordance with the lease agreements before the premises are handed over to the tenant.
> Check that incoming tenant fit-out and outgoing tenant make-good activities are monitored, with particular attention to the relocation of tenants during the period.

Group 2-6 lines under each key process; order: control environment → operating cycle in sequence → receivables / collections → reconciliations.

## Walkthrough background
One paragraph on the business and the process (what, where, which system, who — team sizes matter for segregation of duties), one on changes in the period, one on prior-year points the walkthrough must confirm. The library gives the first; the engagement adds the other two through `background_additions`.
> {CLIENT} ("{CLIENT_SHORT}") operates the {PROCESS} process through the {DEPARTMENT} department, which is headed by the {HEAD_OF_DEPARTMENT}. {FINANCE} records the transactions, performs the reconciliations and reports the results, and {IT} administers access to {SYSTEM}.
> In the current period the company relocated its operations, and the leases at the old premises ended in the same month as the new leases commenced. These two populations are the focus of the tests of new and closed leases.

## Activity and control (walkthrough rows)
- Activity: who does what, when, in which system, with which document — present tense, one activity per row.
> When a unit becomes vacant or new units become available, the {MANAGER} compiles a list of target companies against financial, non-financial and other criteria, and the {HEAD_OF_DEPARTMENT} selects the companies to approach.
- Control: what prevents or detects the risk and how it is evidenced. Thresholds and limits belong to the engagement, not to the library: write "in accordance with the Delegation of Authority" in the library and put the actual limits in the spec.
> The Delegation of Authority is approved by the {CEO} or the Board, states the approving authority and the monetary limits for every transaction type, and is reviewed when the organisation or the limits change.
- Performed by / Control owner: role titles, never names. Frequency: `As and when required`, `Daily`, `Weekly`, `Monthly`, `Quarterly`, `Annually`, `One-off`. Classification `Preventive` / `Detective`; Type `Manual` / `IT` / `Manual and IT`.
- In library text, write the article before a role token (`the {FINANCE_PREPARER} prepares …`) and never start a sentence with a lower-case token; `compose` capitalises sentence starts and removes doubled articles, but clean source reads better.
- Walkthrough document: the exact evidence to sight (`Access request forms and the joiner, mover and leaver listing`).
- A row may cross-refer instead of describing (`Please refer to <A302> for the segregation of duties matrix.`).

## Walkthrough questions
Open questions that make the interviewee describe the step; the "why" states the risk, the prior-year point or the change; documents are specific.
> Q: Could you walk us through the current Delegation of Authority for each type of transaction in scope, including pricing above and below budget, the selection of tenants, the signing of agreements, variations, concessions and waivers, credit notes, refunds and write-offs? Has it changed since the last review?
> Why: We need to confirm the current thresholds and approvers for every transaction type in scope, because the tests of transactions check approvals against these limits.
> Documents: The approved Delegation of Authority and the evidence of its approval.
Interviewee: role. Response: blank (completed in the session). Control ref: the control the question tests. In the library the "why" stays generic ("a frequent prior-year finding is followed up here"); the engagement's actual observation goes into the spec.

## Risk (RCM)
"X may not / may be … This may result in …" — name the specific failure **and** the consequence; mention structural causes (small team, system migration, relocation) where the engagement has them.
> Users may hold access rights in {SYSTEM} that are excessive or incompatible with their roles. Shared accounts and the accounts of resigned staff may remain active, and the activities of the system administrator and external vendor accounts may not be monitored. Unauthorised transactions may therefore not be detected.

## Finding for design effectiveness (planning stage)
`To be concluded after the walkthrough.` followed, where relevant, by the prior-year observation and management's commitment (from `prior_year` in the selection):
> In the prior year we raised Observation 2.2 (Moderate). Three of the five new leases sampled had no insurance policy on file. Management committed to track insurance submissions and coverage in the system.

## Audit strategy / audit steps
Header line `Test n: title`, then the selection sentence or `Check that:` and numbered steps; each step names the attribute and the document that proves it. Where a risk is covered by another test: `Refer to Tests 4 and 5.` or `Refer to Test 1.` followed by the specific step numbers.
> Test 6: Expired and terminated leases, make-good and security deposit refunds
> Select ten expired or terminated leases, or all such leases if the population is smaller, and check that:
> 1. The notice of expiry or termination complied with the {AGREEMENT}, and any early termination penalty was imposed or its waiver approved.
> 2. The joint inspection checklist was signed, and the make-good works were completed or the cost was assessed.

## Test sheet fields
- Objective: `To check that …` (one sentence; add `This test also follows up on prior-year Observation n.` in the spec where it does).
- Basis of selection: `Haphazard, with emphasis on …` / `All …` / `Not applicable, as …`. Name what the emphasis is (largest, non-standard, transition month).
- Population: what and per which listing (`Invoices issued during the period per the invoice listing`).
- Sample size: `Twenty-five across three months` / `Ten, or all new leases if the population is smaller` / `All users and three months`.
- Source: documents and listings used.
- Detail columns: the identifiers and attributes the steps need. Computed columns get a formula (`{r}` = row).
- Notes / Findings / Conclusion defaults: `Nil` / `Nil` / `To be concluded on completion of fieldwork.`

## DA plan rows
Analytic = what is compared with what; source = listings; identifies = what an exception looks like; feeds = the test(s).
> Security deposits held compared with {DEPOSIT_MONTHS} months of rent for each lease. | Lease listing and security deposit listing | A deposit below {DEPOSIT_MONTHS} months of rent indicates a shortfall. | test.cl_deposit_listing, test.cl_closures

## Sizing conventions seen in practice
New / renewed / terminated items: 5-10 (or all if fewer); transactions: 25 across three months; monthly controls: 3 months including the period end or a transition month; all users for access reviews; full listing reconciled + 10 vouched for balances.

## What never goes into a library
Client or group names, tenant or vendor names, product or system brand names (use `{SYSTEM}`, `{ERP}`), people's names or initials, monetary amounts and thresholds, dates and periods, observation numbers and ratings, terminal / site / branch names, anything a reader could trace to one engagement. If the item only makes sense with those details, it belongs in the engagement spec, not in the library. A threshold or duration that the process needs ("items aged beyond N months", "within N days") becomes a token whose default names the policy that sets it (`{GRNI_AGE}` = "the period set in the closing policy"); the engagement supplies the number.
