---
name: emir-validate
description: >
  Validates an EMIR Refit trade CSV against the functional rules in
  reference_data/validation_rules.json, resolving GLEIF/DSB referential
  checks against the mock reference_data JSON files, and produces a
  multi-sheet Excel results workbook (per-trade summary, failures/warnings
  detail, per-rule coverage, full audit trail). Use whenever the user asks
  to validate, test, or check trade data against the EMIR rules, re-run
  validation after editing trade_data_200.csv or validation_rules.json,
  or wants a fresh test_results.xlsx. Trigger for phrasing like "validate
  the trades", "run the EMIR validation", "check this file against the
  rules", "regenerate test results", even without the word "skill".
allowed-tools:
  - Read
  - Bash(python *)
  - Bash(pip install *)
---

# EMIR Refit Trade Validation

Runs the project's rule engine (`scripts/validate_and_report.py`) against a
trade CSV and writes an Excel results workbook. The rule logic is a checked-in
script, not something to re-derive from the conversation each time — read it
before changing behavior, don't reimplement it inline.

## Running it

The trade file to validate is a parameter: if the user names one (as the
skill's argument, e.g. `/emir-validate trade_data_200_v2.csv`, or anywhere
in their request, e.g. "validate trade_data_200_v2.csv"), pass it via
`--trades`. Otherwise omit the flag and the script defaults to
`trade_data_200.csv`.

```
python .claude/skills/emir-validate/scripts/validate_and_report.py --trades <name-or-path>
```

A bare filename (no directory) is resolved against the project root even if
the current shell's working directory differs — this matters when running
from the global copy of this skill (`~/.claude/skills/emir-validate/`)
rather than the project-local one, since the project root is no longer
derivable from the script's own location. You don't need to pass an
absolute path for the common case; just pass the name the user gave you.

Other defaults (override with flags if the user names different files):

| Flag | Default |
|---|---|
| `--rules` | `reference_data/validation_rules.json` |
| `--gleif` | `reference_data/gleif_lei_reference.json` |
| `--dsb` | `reference_data/dsb_upi_reference.json` |
| `--out` | `test_results.xlsx` |

If `openpyxl` is missing, install it first: `pip install openpyxl`
(user-level install, no venv needed for this project).

The script prints a one-line summary (PASS/FAIL/WARNING/N/A counts, trades
with at least one ERROR) — relay that to the user, then point them at the
sheet most relevant to what they asked for (see below). Don't dump the raw
per-row output into the response; the workbook is the deliverable.

## Output workbook

- **Trade Summary** — one row per trade: overall status (FAIL if any ERROR,
  else WARNING if any WARNING, else PASS), counts, and the specific rule IDs
  that fired. Start here for "which trades failed".
- **Failures And Warnings** — every FAIL/WARNING finding with the full rule
  message. Start here for "why did trade X fail" or "show me every R0NN
  violation".
- **Rule Summary** — PASS/FAIL/WARNING/N/A counts per rule across the whole
  file. Use this to sanity-check that a rule change actually shifted
  behavior, or that a fixture still exercises every rule.
- **All Results** — full audit trail, every (trade, rule) pair, including
  PASS and N/A. Only pull from this sheet for exhaustive/regression asks.

## What each rule `type` needs to be evaluated (for when rules change)

`reference_data/validation_rules.json` entries have a `type` field that
determines evaluation scope. If the user adds or edits a rule, match its
`condition` text to the right scope before writing the evaluator function:

| Type | Scope | Engine mechanism |
|---|---|---|
| `Format` | single field | regex on `EVALUATORS["Rxxx"]` |
| `Mandatory` | single field | not-blank check |
| `Range` | single field | numeric/date bound |
| `Consistency` | same row (or same-`UTI` group for lifecycle rules) | reads other columns on `row`, or `self.by_uti[row["UTI"]]` |
| `Uniqueness` | across all rows | `self.newt_by_generator` grouping |
| `Referential` | external lookup | `self.gleif` / `self.dsb` dicts, keyed by `lei` / `upi` |
| `Timeliness` | timestamp vs timestamp+deadline | datetime arithmetic |
| `Plausibility` / `DataQuality` | soft, `WARNING`-severity only | same mechanics as above, just non-blocking |

New rule IDs must follow the `Rnnn` (3-digit, zero-padded) convention used
throughout — the engine derives `EVALUATORS` from `r001`..`r050` by name, so
add a `def rNNN(self, row):` method with a matching ID and extend the range
in `Engine.__init__` if the rule count grows past 50.

## Known limitations — don't silently "fix" these, they're documented gaps

- **`R039`** (COMP requires ≥2 linked UTIs) always returns `N/A` — the CSV
  schema has no `LinkedUTIs` column. Don't fabricate a pass/fail for it;
  extending the schema is a separate task.
- **`R021`** (UTI must not change across a lifecycle) links rows via a
  TradeID-suffix heuristic (`lifecycle_key()`), not a real regulatory field.
  It only works because this fixture's TradeIDs are tagged
  `TRD-Rnnn-...-<n>-<ACTIONTYPE>`. It will not link lifecycle events in
  real-world data — that needs an explicit `PriorUTI`/lifecycle-link column.
- **`R033`**'s notion of "the submitting entity" is the hardcoded
  `SUBMITTING_ENTITY_LEI` constant at the top of the script, not something
  derivable from the trade file. Update that constant (or make it a CLI flag)
  before validating a trade file for a different reporting entity.
- **`R022`/`R023`/`R024`/`R042`/`R043`/`R044`** are only as trustworthy as
  `reference_data/gleif_lei_reference.json` and `dsb_upi_reference.json`.
  These are mocks; an LEI or UPI absent from them is treated as not-found and
  fails/warns accordingly. Point `--gleif`/`--dsb` at richer files if the
  user wants to validate LEIs/UPIs beyond the ones seeded there.

Full narrative context (file map, CSV `TradeID` naming convention, why the
package looks the way it does) lives in `README.md` at the project root —
read it if any of the above needs more background than this file gives.
