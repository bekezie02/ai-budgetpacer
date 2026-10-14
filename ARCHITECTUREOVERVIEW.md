# Architecture Overview — AI BudgetPacer

For newcomers to the codebase. See [README.md](./README.md) for a high-level summary and [CONTRIBUTIONGUIDE.md](./CONTRIBUTIONGUIDE.md) for setup.

## Core files
| File | Responsibility |
|---|---|
| `financeUtils.gs` | Plaid Link handshake + `transactions/sync` fetch (cursor-paginated); also hosts the shared `llm()` wrapper (used by this project and Toolsmith Agent) |
| `financeCore.gs` | Categorization, ledger, fixed pools, sheet updates, sync orchestration, Unknown bucket |
| `financeEmail.gs` | Builds the snapshot, generates the LLM summary, sends the daily email |
| `financeSchedule.gs` | Triggers — daily sync/email, and a 15-minute cron for checking email replies |

## Spreadsheet layout
- **Annual Overview** — the source of truth for three annual targets:
  - `Total Mandatory Costs Excluding Fixed Amounts` (non-discretionary, excluding the individually-tracked fixed categories)
  - A block of individually-tracked **fixed categories** (Renters Insurance, Medical, Car Maintenance, Car Registration, Car Insurance), each with its own budget column, plus Travel with its own total line
  - `VARIABLE / LIFESTYLE COSTS (ANNUAL ESTIMATE)` ending in `Total Variable Costs`
  - `Projected Annual Savings` (the annual savings target)
- **Monthly LifeStyle Cost** / **Monthly Mandatory Cost** sheets — one section per month, located by scanning column A (layout is irregular, not fixed row numbers). Each month's `Total Variable Costs` row has Amount/actual (col B), Budget/expected (col C), and a computed Savings/difference (col D). Year-to-date rollups live in fixed cells (Lifestyle: F2 actual / G2 savings; Mandatory: H2 actual / I2 savings), recomputed automatically.
- **Fixed pool sheets** (one "Fixed - <label>" pool per fixed category + Travel) — each has an Annual Budget and a Remaining Balance.
- **Uncategorized Transactions ("Unknown") sheet** — 9 columns (including merchant_name, pfc_primary, pfc_detailed, and an assignment dropdown) holding transactions that didn't match any Category Rule.
- **Category Rules sheet** — Merchant Name Contains | Plaid Category (primary) Contains | Plaid Category (detailed) | Target Sheet | Target Field, AND-matched.
- **Transaction Ledger** — source of truth for applied target and per-transaction category overrides. The Override Category dropdown can move an existing transaction to another monthly field, Fixed pool, or Ignore without changing its reusable Category Rule; choosing "Use original category" clears the override.

## How a day flows
1. `dailyFinanceCheck()` fires (daily trigger).
2. `ensureMonthSectionsUpToDate()` runs first — builds any missing month (Jan → current) and reconciles every existing month's categories/budgets against the Annual Overview blocks (never touching Actual).
3. Plaid transactions are synced (`transactions/sync`, cursor-paginated, page size 50) and categorized via `applyTransaction()` against Category Rules; matches update the ledger, `"Ignore"` matches are dropped, and non-matches go to the Unknown sheet.
4. Annual/monthly rollups and fixed-pool numbers are recomputed (once per sync, not per-transaction, for performance).
5. `buildFinanceSnapshot()` assembles Lifestyle, Mandatory, and Fixed-pool numbers plus day-of-year pacing (`getYearProgressPercent()`).
6. `generateFinanceEmailSummary()` sends the snapshot to the LLM (via `llm()`, `reasoningEffort: "low"`) to write the qualitative narrative — with a retry-once-then-plain-fallback if the response is truncated, and a `hasAnyDataThisYear` guard so it never fabricates pacing commentary from an empty month.
7. The email (HTML) is sent; if it includes Unknown-bucket entries, `dailyFinanceCheck()` stores that Gmail thread ID so replies can be tracked.
8. The summary email lists unassigned transactions and links directly to the Uncategorized Transactions sheet. Selecting a category in its dropdown calls `resolveUnknownTransaction()` — which also assigns other Unknown transactions with the same merchant + PFC primary + PFC detailed, and auto-creates a matching Category Rule. A sheet toast confirms the assignment; no follow-up email is sent for dropdown assignments.
9. Email replies remain supported: `checkForCategoryReplies()` runs every 15 minutes, reads unread replies on the stored thread (ignoring quoted content), asks the LLM to match reply text to an exact category whitelist, and calls the same `resolveUnknownTransaction()` flow.

## Known Groq/LLM quirk
`openai/gpt-oss-120b` (via Groq) is a reasoning model whose hidden chain-of-thought tokens count against `max_completion_tokens` even when not shown — a fixed 1,200-token budget at "medium" reasoning effort could get consumed before the visible answer finished. `llm()` now accepts an optional `{maxCompletionTokens, reasoningEffort}` params object (defaulting to prior behavior), and both finance-email LLM calls pass `reasoningEffort: "low"`.

## Performance/reliability notes for contributors
- Rollup recompute (`recomputeAnnualRollup()`, a full-sheet scan) is intentionally run once per sync/reply, not per-transaction — running it per-transaction previously caused large backlog syncs to hit Apps Script's 6-minute execution limit before the cursor saved, stalling sync at the same point every run.
- A 5-minute time-budget guard in `runDailyFinanceSync` checkpoints the cursor between pages as extra protection.
- The daily sync trigger also installs a one-minute continuation trigger. Initial backfills resume from per-Item cursors until every active Item is caught up, then the script emails a completion notice (and a finance summary if data changed).
- Plaid Item replacements use `startPlaidItemReplacementLink()`, automatic resumable staging through the one-minute inbox trigger (`stagePendingPlaidItemTransactions()` remains available for manual recovery), and `finalizePlaidItemReplacement()` so a new connection's transactions are reviewed before the active Item ID/token is swapped. Staging pre-fills `REPLACE <old transaction ID>` for a unique exact date/amount/description match and `IMPORT` otherwise. Multiple possible matches are flagged for review; `REPLACE` removes the existing Ledger/Unknown transaction before applying the new Item's ID, `IMPORT` intentionally keeps both, and `SKIP` ignores the staged row.
- The replacement review sheet's old-Item dropdown follows `PLAID_ITEM_IDS` order and shows bank name plus Item ID; bank names are looked up from Plaid and cached in Script Properties.
- Transaction Ledger edits to Override Category reclassify only that transaction, adjusting the previous and destination budgets and preserving the selected override through later Plaid updates. The adjacent default-target columns retain its original category so the override can be cleared; `setupFinanceSheets()` installs the required edit trigger and dropdown.
- Replacement finalization is also time-bounded and resumable: each completed row is marked before the next batch, and the one-minute inbox trigger continues after the single `FINALIZE <item ID>` approval until it completes.
- On successful finalization, `PLAID_ITEM_IDS` is switched to the new Item; its Item-keyed access token and staged sync cursor remain for normal sync, the old Item's token/cursor are removed, and the temporary review sheet is deleted.
- Date handling: Sheets silently converts a written `"YYYY-MM-DD"` string into a real Date object, which can break re-deriving a transaction's month. `getMonthYearFromDate()` handles both string and Date input, with guards that return a clear error instead of crashing.
