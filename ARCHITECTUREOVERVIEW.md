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
- **Uncategorized Transactions ("Unknown") sheet** — 8 columns (including merchant_name, pfc_primary, pfc_detailed) holding transactions that didn't match any Category Rule.
- **Category Rules sheet** — Merchant Name Contains | Plaid Category (primary) Contains | Plaid Category (detailed) | Target Sheet | Target Field, AND-matched.

## How a day flows
1. `dailyFinanceCheck()` fires (daily trigger).
2. `ensureMonthSectionsUpToDate()` runs first — builds any missing month (Jan → current) and reconciles every existing month's categories/budgets against the Annual Overview blocks (never touching Actual).
3. Plaid transactions are synced (`transactions/sync`, cursor-paginated, page size 50) and categorized via `applyTransaction()` against Category Rules; matches update the ledger, `"Ignore"` matches are dropped, and non-matches go to the Unknown sheet.
4. Annual/monthly rollups and fixed-pool numbers are recomputed (once per sync, not per-transaction, for performance).
5. `buildFinanceSnapshot()` assembles Lifestyle, Mandatory, and Fixed-pool numbers plus day-of-year pacing (`getYearProgressPercent()`).
6. `generateFinanceEmailSummary()` sends the snapshot to the LLM (via `llm()`, `reasoningEffort: "low"`) to write the qualitative narrative — with a retry-once-then-plain-fallback if the response is truncated, and a `hasAnyDataThisYear` guard so it never fabricates pacing commentary from an empty month.
7. The email (HTML) is sent; if it includes Unknown-bucket entries, `dailyFinanceCheck()` stores that Gmail thread ID so replies can be tracked.
8. Separately, `checkForCategoryReplies()` runs every 15 minutes, reads unread replies on the stored thread (ignoring quoted content), asks the LLM to match reply text to an exact category whitelist, and calls `resolveUnknownTransaction()` — which also sweeps and auto-resolves any other Unknown transaction with the same merchant + same PFC-primary, and auto-creates a new Category Rule from the resolution.

## Known Groq/LLM quirk
`openai/gpt-oss-120b` (via Groq) is a reasoning model whose hidden chain-of-thought tokens count against `max_completion_tokens` even when not shown — a fixed 1,200-token budget at "medium" reasoning effort could get consumed before the visible answer finished. `llm()` now accepts an optional `{maxCompletionTokens, reasoningEffort}` params object (defaulting to prior behavior), and both finance-email LLM calls pass `reasoningEffort: "low"`.

## Performance/reliability notes for contributors
- Rollup recompute (`recomputeAnnualRollup()`, a full-sheet scan) is intentionally run once per sync/reply, not per-transaction — running it per-transaction previously caused large backlog syncs to hit Apps Script's 6-minute execution limit before the cursor saved, stalling sync at the same point every run.
- A 5-minute time-budget guard in `runDailyFinanceSync` checkpoints the cursor between pages as extra protection.
- Date handling: Sheets silently converts a written `"YYYY-MM-DD"` string into a real Date object, which can break re-deriving a transaction's month. `getMonthYearFromDate()` handles both string and Date input, with guards that return a clear error instead of crashing.
