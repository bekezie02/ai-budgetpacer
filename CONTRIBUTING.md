# Contributing — AI BudgetPacer

Coding standards and submission steps for this project. See [README.md](./README.md) for what it does and how to install it, and [ARCHITECTUREOVERVIEW.md](./ARCHITECTUREOVERVIEW.md) for how the pieces fit together.

## Coding standards

- **Deterministic core, LLM only at the edges.** Any new feature that can be computed in code (rollups, percentages, pacing, matching) must be — route new logic through `llm()` only for genuinely qualitative writing (the email summary) or free-form-text parsing (reply parsing). Don't let the LLM own arithmetic or category-matching logic.
- **Category Rules stay AND-matched.** If you add a new match dimension (as was done for `pfc_detailed`), it must AND with existing columns, not replace or OR with them, and you must document the required manual migration for existing rows in this file's migration notes below.
- **Respect the 6-minute execution limit.** Never add per-transaction work inside the sync loop (e.g. a full-sheet rollup recompute) — batch it to once-per-sync/once-per-reply, following the existing `recomputeAnnualRollup()` pattern, and keep the sync's time-budget guard and cursor checkpointing intact.
- **Guard against silent Sheets type coercion.** Any function reading a date-like cell should accept both string and real `Date` values (see `getMonthYearFromDate()`) and fail loudly rather than silently miscomputing.
- **Never fabricate progress commentary.** LLM-generated summary text must respect the `hasAnyDataThisYear` (or equivalent) guard — no pacing/progress narrative from empty data.
- **Secrets never in source.** Plaid and Groq credentials belong in Script Properties, never hardcoded or committed.
- **Respect file boundaries:**
  - `financeUtils.gs` — Plaid handshake/sync + shared `llm()` wrapper only.
  - `financeCore.gs` — categorization, ledger, fixed pools, Unknown bucket.
  - `financeEmail.gs` — snapshot assembly, LLM summary, sending.
  - `financeSchedule.gs` — triggers only.

## Submission steps

1. **Branch from the latest pushed state.** `clasp pull` first if the live Apps Script project may have changed via the browser IDE since your last pull; then create a git branch off `main`.
2. **Implement the change**, respecting the file boundaries and standards above.
3. **Test against a non-production spreadsheet or Plaid sandbox** before pointing at real data — this project touches live financial data and a live Plaid connection, so avoid testing destructive changes (e.g. category rule matching, sweep-resolution logic) directly on the production sheet.
4. **Manually trigger the affected flow** (`dailyFinanceCheck()`, `checkForCategoryReplies()`, etc.) and confirm: no full-sheet operations were added inside a per-transaction loop, dates parse correctly, and the Unknown-sweep / auto-rule-creation behavior still matches on the intended fields only.
5. **Migrate existing sheet data if your change alters a schema** (e.g. Category Rules columns, fixed pool structure) — note the required manual migration steps in this file and in the README's "Known limitations" section.
6. **Update `docs/SHEET_LAYOUT.md`** if the change alters the spreadsheet schema, so the public repo's structure snapshot stays accurate — it's a point-in-time export, not synced automatically.
7. **`clasp push`** to sync, then do one more manual run against test data to confirm the deployed version behaves as tested.
8. **Commit and open a pull request on GitHub.** Describe which sheets/columns it touches, whether it requires manual data migration, and how you tested it (sandbox vs. production, which flow you triggered). Never include real transaction data or account numbers in commits, PR descriptions, or screenshots. Merge only after the manual test in step 4/7 passes.
