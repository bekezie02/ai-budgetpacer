# AI BudgetPacer (Google Apps Script + Plaid)

**AI BudgetPacer** is a Google Apps Script personal finance tracker that pulls transactions via Plaid, categorizes and ledgers them against an annual budget, and emails a daily deterministic + LLM-narrated summary.

See also: [Project Changes Guide](./PROJECT_CHANGES_GUIDE.md) · [ARCHITECTUREOVERVIEW.md](./ARCHITECTUREOVERVIEW.md) · [CONTRIBUTING.md](./CONTRIBUTING.md)

## What this is
A daily-cron finance tracker: it syncs bank transactions from Plaid, categorizes them against user-defined rules, rolls them up into annual/monthly budget sheets, and emails a summary. Unrecognized transactions go into an "Unknown" bucket. Assign them from the sheet's category dropdown; the system learns a rule for matching future transactions.

## Architecture philosophy
- **Deterministic core, LLM only at the edges.** There is no LLM tool-orchestration/agent loop here. The LLM is used one-shot to write the qualitative email summary/feedback.
- All percentages, rollups, and pacing math are computed in code — never left to the LLM.

## Notable design choices
- **401k** is pre-tax/invisible to Plaid and intentionally not tracked.
- **Brokerage/Roth IRA** transfers could be tracked but deliberately aren't — no budget tension to reconcile, just a known fixed transfer.
- **Category Rules** are AND-matched across three Plaid fields (Merchant Name Contains, Plaid Category primary, Plaid Category detailed) plus a Target Sheet/Field, so the same merchant can route differently depending on context (e.g. Target for groceries vs. office supplies), and loan-payment types (credit card vs. mortgage vs. student/personal loan) can be told apart.
- For fixed pools, use `FixedPool:<pool name>` as the Category Rules target (for example, `FixedPool:Travel`). Existing rules or Ledger entries using a pool's bare name (for example, `Travel`) are recognized as that pool when its `Fixed - Travel` sheet exists.
- A rule's Target Sheet can be `"Ignore"` — matching transactions (e.g. credit card payments, already counted when the original purchase posted) are skipped entirely: not tracked, not ledgered, not sent to Unknown.
- Resolving an Unknown transaction auto-creates a matching Category Rule (if one doesn't already exist), so repeat merchants can be categorized automatically.
- The finance email lists unassigned transactions and links directly to the Uncategorized Transactions sheet. Choose a category in that row's **Assign to Category** dropdown; matching unassigned transactions are assigned together, and a sheet confirmation reports the result.
- To reclassify one already-budgeted transaction without changing future transactions, select a value in that transaction's **Override Category** column in the Transaction Ledger. For example, move one dining transaction to **Fixed — Travel** while leaving the usual eating-out rule unchanged. Select **Use original category** to clear the override.

## Known limitations / decisions
- Only Jan–current-month sections exist in the Monthly sheets at any time (built up to the current month, not pre-built for the full year) — this is intentional, not a bug.
- The old single combined "Fixed - Medical and Car" pool is now orphaned; any historical spend in it needs manual migration into the newer individually-tracked fixed pools (Renters Insurance, Medical, Car Maintenance, Car Registration, Car Insurance, Travel).

## Repository structure
This repo is a **public mirror** of a live Google Apps Script project, bridged via [`clasp`](https://github.com/google/clasp) — Apps Script has no local runtime, so the `.gs` files here are the same source that runs in Google's cloud, not a separate local version.

```
.
├── financeUtils.gs           # Plaid handshake/sync + shared llm() wrapper
├── financeCore.gs            # categorization, ledger, fixed pools, Unknown bucket
├── financeEmail.gs           # snapshot assembly, LLM summary, sending
├── financeSchedule.gs        # triggers
├── appsscript.json           # Apps Script manifest (pulled by clasp)
└── docs/
    └── SHEET_LAYOUT.md        # snapshot describing the spreadsheet structure (Annual Overview, Monthly sheets, Fixed pools, Category Rules)
```
This project's "documents" aren't Google Docs like the agent project's — they're the spreadsheet's own structure (sheet names, columns, anchor rows). `docs/SHEET_LAYOUT.md` is a **point-in-time export** of that structure for public readability, since the live spreadsheet itself (with real transaction data) obviously isn't published. Re-export it after any schema change.

## Required properties / environment variables
This project runs on Google Apps Script, so configuration lives in **Script Properties** (Project Settings → Script Properties) rather than a `.env` file — none of this lives in the repo.

| Property | Purpose | Required |
|---|---|---|
| `PLAID_CLIENT_ID` | Plaid API auth | Yes |
| `PLAID_SECRET` | Plaid API auth | Yes |
| `PLAID_ENV` (e.g. `sandbox`/`development`/`production`) | Which Plaid environment to hit | Yes |
| `GROQ_API_KEY` | Auth for the shared `llm()` (summary writing) | Yes |
| `FINANCE_SPREADSHEET_NAME` | Points the script at the active finance spreadsheet; repointed by `startFreshFinanceSheet()` | Yes |

> Note: `SYNC_START_DATE` is calculated in code and is not a Script Property. The Plaid/Groq credential names are conventional and should be confirmed against your live Script Properties. `runStep1()` calculates the number of days from January 1 through today when creating a Plaid Link token, so the current date is never hardcoded.

## Deploy your own copy

> Apps Script runs on Google's servers, not on your machine — "deploying" here means pushing this repo's code into your own Apps Script project via `clasp`, not running it locally. Your actual financial data always stays in your own private spreadsheet, never in this repo.

### Prerequisites
- Node.js (for `clasp`) and a Google account with Apps Script + Sheets + Gmail access.
- A Plaid developer account with API credentials (client ID + secret) and at least one linked bank connection.
- A Groq API key (shared setup with Toolsmith Agent).

### Steps
1. **Install clasp** and log in: `npm install -g @google/clasp` then `clasp login`.
2. **Clone this repo**, then either `clasp create` (new Apps Script project) or point `.clasp.json` at an existing `scriptId` you own.
3. **Set Script Properties** in the Apps Script UI per the table above — never commit Plaid or Groq credentials to the repo.
4. **Set up your own spreadsheet** using `docs/SHEET_LAYOUT.md` as the reference: Annual Overview (with the Mandatory/Lifestyle/fixed-category/Travel blocks and Projected Annual Savings), Monthly LifeStyle Cost, Monthly Mandatory Cost, one Fixed pool sheet per fixed category + Travel, Uncategorized Transactions, Category Rules.
5. **Link bank accounts via Plaid Link** and complete the handshake in `financeUtils.gs`.
6. **Push and deploy**: `clasp push` to sync the code into your Apps Script project.
7. **Run `setupFinanceSheets()`** (creates the Unknown sheet and any other structural setup) if starting fresh. After deploying updates to an existing spreadsheet, run it again to add the Transaction Ledger override columns, Uncategorized Transactions dropdown, and edit trigger.
8. **Wire up triggers**: a one-minute trigger for `checkInboxForTriggers()` (install with `setFinancialTrigger()`) and daily sync plus one-minute backfill-continuation triggers (installed with `setDailyFinanceTrigger()`). `setupFinanceSheets()` also removes any obsolete email-category trigger.
9. **Populate initial Category Rules** for known recurring merchants, including any `"Ignore"` rules (e.g. credit card payment transactions).
10. **First run / backfill**: run `dailyFinanceCheck()` once. If the Plaid history takes longer than one execution, its saved cursors resume automatically each minute. The script sends a caught-up email when all active Items reach the latest available updates; no repeated manual run is needed.

### Keeping the repo and live script in sync
Pick one direction as source of truth to avoid silently overwriting changes:
- **Repo is source of truth**: edit `.gs` files locally, `clasp push` to deploy.
- **Apps Script UI is source of truth**: edit in the browser editor, `clasp pull` before your next local edit, then commit.

### Starting a new year / fresh sheet
Use `startFreshFinanceSheet(newSpreadsheetName)` — it repoints `FINANCE_SPREADSHEET_NAME` at the new sheet and clears sync cursors (bounded by `SYNC_START_DATE`) while keeping the same linked Plaid bank connections.

### Re-syncing transactions from the beginning
To replay transactions for the current spreadsheet, run `resyncTransactionsFromStart()` manually in the Apps Script editor. It clears only the linked Plaid items' saved sync cursors and immediately starts the normal paginated sync; it does not clear spreadsheet data. Existing Ledger and Unknown entries are matched by Plaid transaction ID, so they are updated rather than duplicated. If the sync reaches its execution time limit, the saved cursors let subsequent regular syncs continue. The replay is still limited to the transaction history Plaid makes available for each linked account.

### Replacing a Plaid Item with a new Link connection
Replacements are handled one at a time; after finalizing or cancelling one, you can start another. Each new Item gets its own `Plaid Review <item ID>` sheet, so later replacements do not reuse an earlier Item's staging data. To test a new Link connection with the current-year history request without immediately syncing it into the budget:
1. Send yourself an email with the subject `Plaid Item Replacement`. The existing one-minute `checkInboxForTriggers` schedule emails you the Hosted Link URL and checks for Link completion on subsequent trigger runs for up to 25 minutes (without keeping an execution sleeping).
2. The new Item is held out of normal syncing. After the handshake, the one-minute inbox trigger emails you when transaction staging starts, automatically stages transactions in resumable batches, and emails you the review sheet when staging is complete.
3. If staging encounters an error, the script emails the error and retries on a later trigger run. `stagePendingPlaidItemTransactions()` remains available for manual recovery; if it reports `in_progress`, run it again. The review sheet automatically adds rows as needed.
4. In the generated `Plaid Review <item ID>` sheet, choose the old Item in cell B1. The dropdown follows the order in `PLAID_ITEM_IDS` and shows each institution name alongside its Item ID. Decisions are prefilled: exactly one match on date, amount, and normalized description gets `REPLACE <old transaction ID>`; all other transactions default to `IMPORT`. Change `IMPORT` to `SKIP` for any transaction you do not want. Multiple possible matches are flagged with status `REVIEW` and shown in column K; check those carefully because the default `IMPORT` could add a duplicate. `REPLACE` removes the old Ledger/Unknown entry before applying the new Item's transaction ID; `IMPORT` deliberately keeps both transactions; `SKIP` does not import the staged row.
5. After every transaction has a decision, use the review sheet's **Replacement action** dropdown in B5 to finalize. Finalization processes rows in time-bounded batches, records each completed row's status, and automatically resumes on later one-minute trigger runs. It imports/replaces only rows marked `IMPORT` or `REPLACE`, swaps the old Item ID in `PLAID_ITEM_IDS` for the new one, keeps the new Item's `PLAID_ACCESS_TOKEN_<item ID>` and staged `PLAID_CURSOR_<item ID>`, removes the old Item's token and cursor properties, and deletes the completed replacement's review sheet. You receive a completion email.

To cancel a staged replacement and start over, choose **Cancel replacement** in the review sheet's B5 dropdown. It cancels an in-progress replacement Link session or a staged replacement, removes the pending Item's local token/cursor, and deletes that replacement's review sheet. It refuses to delete a review sheet after any transaction has been imported or replaced. Cancellation does not remove the Item from Plaid or change the active Item or finance data. You can also run `cancelPlaidItemReplacement()` directly. The requested history depends on the bank's availability. Possible duplicates are suggestions only because transaction IDs from separate Items are not guaranteed to match; verify each suggested match before choosing `REPLACE`.

Unassigned transactions can only be assigned from the Uncategorized Transactions sheet's dropdown; replies to the finance email are not processed as category assignments.

Automatically growing sheets (monthly sections, fixed-pool ledgers, Transaction Ledger, Uncategorized Transactions, Category Rules, and Plaid review sheets) extend their row capacity as needed; manual row expansion should not be necessary.
