# EMIR REFIT Test Package

A self-contained fixture set for testing an EMIR (Regulation (EU) No 648/2012, as
amended by EMIR Refit) trade-report validator: a set of ESMA-style functional
validation rules, a matching set of test trades engineered to hit every pass /
fail / warning path in those rules, and the mock reference data those rules
depend on.

It is designed to be consumed by a validation engine (script, service, or
Claude Code skill) that needs to prove out its rule coverage before pointing
at a real trade repository feed.

## Contents

| File | Purpose |
|---|---|
| `reference_data/validation_rules.json` | The 50 functional validation rules under test. |
| `trade_data_200.csv` | 200 test trade records, purpose-built to exercise every rule. |
| `reference_data/gleif_lei_reference.json` | Mock GLEIF LEI-status lookup (backs the LEI referential rules). |
| `reference_data/dsb_upi_reference.json` | Mock DSB UPI-status lookup (backs the UPI referential rules). |
| `expected_results.csv` | Expected outcome per `TradeID`, for regression comparison. **Currently stale** — see [Known limitations](#known-limitations--gaps). |
| `.claude/skills/emir-validate/` | The validation engine as a Claude Code skill — see [Validation skill](#validation-skill). |
| `test_results.xlsx` | Generated output of the validation skill — not checked in by hand, regenerate via the skill. |

## `validation_rules.json` schema

Each rule is a flat object:

```json
{
  "ruleId": "R038",
  "field": "ReportingTimestamp",
  "type": "Timeliness",
  "condition": "ReportingTimestamp <= ExecutionTimestamp + 1 working day",
  "severity": "ERROR",
  "message": "Report must be submitted no later than the working day following ..."
}
```

- **`field`** — one of the 10 EMIR Refit fields in scope: `UTI`, `UPI`,
  `Counterparty1LEI`, `Counterparty2LEI`, `NotionalAmount`, `Currency`,
  `MaturityDate`, `ReportingTimestamp`, `ActionType`, `AssetClass`.
- **`type`** — the category of check, which determines how an engine should
  evaluate `condition`:
  | Type | Evaluated against | Example |
  |---|---|---|
  | `Format` | the field value alone (regex/pattern) | UTI matches `^[A-Z0-9]{18,52}$` |
  | `Mandatory` | the field value alone (null/blank check) | `UTI IS NOT NULL` |
  | `Range` | the field value alone (numeric/date bounds) | `NotionalAmount > 0` |
  | `Consistency` | multiple fields on the **same row** | `NotionalAmount.currency = Currency` |
  | `Uniqueness` | the field value **across all rows** | no duplicate `UTI` for a given generating LEI |
  | `Referential` | the field value against an **external lookup** | UPI must be `ACTIVE` in the DSB reference data |
  | `Timeliness` | a timestamp field against another timestamp + a deadline | `ReportingTimestamp <= ExecutionTimestamp + 1 working day` |
  | `Plausibility` / `DataQuality` | soft, non-blocking sanity checks | notional > 10bn, LEI renewal due soon |
- **`severity`** — `ERROR` (must block submission) or `WARNING` (should be
  surfaced but not block).
- **`condition`** — a human/engine-readable pseudo-expression, not a specific
  language. Treat it as the functional spec for the check; a real
  implementation compiles it to code, SQL, or a rules-engine expression.

Rules that reference fields lower-cased with a dot (e.g.
`Counterparty1LEI.status`) or that mention "GLEIF"/"DSB" require an external
lookup — see [Reference data files](#reference-data-files).

## `trade_data_200.csv` schema

```
TradeID,UTI,UPI,ActionType,EventDate,ExecutionTimestamp,ProductType,AssetClass,
Counterparty1LEI,Counterparty2LEI,NotionalAmount,Currency,MaturityDate,
ClearingObligation,Cleared,ReportingTimestamp
```

`TradeID` is not itself governed by any rule — it is used purely to make the
file self-documenting:

- **`TRD-R0NN-PASS` / `TRD-R0NN-FAIL` / `TRD-R0NN-WARN`** — a row engineered
  to satisfy or violate rule `RNN` specifically, with every other field kept
  valid so the row isolates that one rule. `WARN`-suffixed rows target
  `WARNING`-severity rules.
- **`TRD-R0NN-...-1-NEWT`, `-2-MODI`, `-2-CORR`, ...** — multi-row sequences
  for rules that can only be evaluated against a derivative's reporting
  history rather than a single row in isolation: `R021` (UTI must not change
  across lifecycle events), `R029` (a MODI/CORR/... must have a prior NEWT
  with the same UTI), `R031` (UTI must be unique per generating LEI), and
  `R049` (more than 3 CORR/EROR reports for the same UTI within 30 days is a
  data-quality warning). An engine evaluating these must process the file as
  a batch, grouped by `UTI`, not row-by-row.
- **`TRD-BASE-####`** — 92 filler rows with varied but fully valid data
  (rotating asset class, currency, counterparty pair, notional). These carry
  no rule-specific intent; they exist as general "everything should pass"
  regression coverage and to pad the fixture to 200 rows.

Row count: 200 (108 targeted + 92 baseline).

## Reference data files

Six rules (`R022`, `R023`, `R024`, `R042`, `R043`, `R044`) can't be evaluated
from `trade_data_200.csv` alone — they require the kind of external lookup a
production validator would get from GLEIF (LEI registry) or the DSB (UPI
registry). These are mocked here so the fixture set is self-contained:

- **`gleif_lei_reference.json`** — keyed by `lei`, drives `R023`/`R024`
  (fails if `registrationStatus` is not `ISSUED` or `LAPSED`) and
  `R043`/`R044` (warns if `nextRenewalDate` is within 30 days of the trade's
  `ReportingTimestamp`, or if status is `LAPSED`). Any LEI in a trade row
  that isn't in this table should be treated as not found and fail
  `R023`/`R024`.
- **`dsb_upi_reference.json`** — keyed by `upi`, drives `R022` (fails if the
  UPI isn't present with `status = ACTIVE`) and `R042` (warns if
  `lastRefreshDate` is more than 180 days before the trade's
  `ReportingTimestamp`). `ZZZUNKNOWN01` is deliberately absent from the table
  — that absence *is* the R022 fail case.

A real integration replaces both files with live calls to GLEIF and the DSB;
the lookup key and the fields consumed (`registrationStatus`,
`nextRenewalDate`, `status`, `lastRefreshDate`) are what a production adapter
needs to supply.

## Known limitations / gaps

- **`expected_results.csv` is stale.** It still lists outcomes for the old
  sequential `TradeID`s (`TRD0001`...`TRD0200`) from before the file was
  regenerated with rule-tagged `TradeID`s (`TRD-R0NN-...`, `TRD-BASE-####`).
  It needs to be regenerated against the current `trade_data_200.csv` before
  it can be used for regression comparison.
- **`R039`** (a `COMP` action must reference at least two prior linked UTIs)
  cannot be tested with the current CSV schema — there is no `LinkedUTIs`
  column. `trade_data_200.csv` includes one illustrative `COMP` row
  (`TRD-R039-PASS-NOTE-SCHEMA-LIMIT`) without asserting pass/fail; adding
  real coverage requires extending the schema.
- Rules typed `Referential` are only as good as the mock reference data —
  they prove the validator's *lookup logic* works, not that it's wired to a
  live GLEIF/DSB feed.

## Validation skill

The evaluation flow described in earlier revisions of this doc is now
implemented as a Claude Code skill: **`.claude/skills/emir-validate/`**. It
loads `reference_data/validation_rules.json`, evaluates every rule against
`trade_data_200.csv` (grouping by `UTI` for lifecycle/uniqueness/frequency
rules, and consulting `gleif_lei_reference.json` / `dsb_upi_reference.json`
for `Referential` rules), applies `ERROR`/`WARNING` severity semantics, and
writes a 4-sheet `test_results.xlsx` (per-trade summary, failures/warnings
detail, per-rule coverage, full audit trail).

Run it directly:

```
python .claude/skills/emir-validate/scripts/validate_and_report.py
```

or ask Claude Code to "validate the trades" / "run the EMIR validation" and
it will invoke the skill. See `.claude/skills/emir-validate/SKILL.md` for
the flag reference, the rule-`type` → evaluation-scope mapping (needed when
adding new rules), and specifics on each known limitation below (which
constant or heuristic each one lives behind in the script).

Comparing against `expected_results.csv` for engine regression testing isn't
wired up yet — that file needs regenerating against the current `TradeID`
scheme first (see below).
