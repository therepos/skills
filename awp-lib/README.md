# awp-lib — shared content libraries for the awp-* skills

`core.json` holds the modules that apply to every process (policies and Delegation of Authority, segregation of duties, system access, master file, reconciliations, variance analysis, general questions). `pack-<process>.json` holds one process (today: `commercial-leasing`). Each file carries its own `version`, `updated` and `changelog`.

A module bundles everything about one sub-process: walkthrough activities with their controls, walkthrough questions, RCM risks, tests and analytics, all by id, with tokens (`{SYSTEM}`, `{HEAD_OF_DEPARTMENT}`, `{REFUND_DAYS}` …) where an engagement supplies the specifics. The full format, the token table and what `validate` checks are in `awp-generate/ref/schema.md`; writing rules are in `awp-generate/ref/content.md`.

Rules:
- **No client data, ever** — no client, tenant, vendor, system, person, amount, date, site or prior-year observation. `python awp-generate/scripts/awp.py check --deny names.txt` must pass before any commit here.
- Content changes arrive through the **awp-learn** skill (`apply` writes the file with its version bump and changelog entry) or from owners by hand, followed by `validate`.
- Every `awp-*` release zip carries a snapshot of this folder as `lib/`; users never read this repository. Release every `awp-*` skill after a library change.
