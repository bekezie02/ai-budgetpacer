# AI BudgetPacer (Google Apps Script + Plaid)

**AI BudgetPacer** is a Google Apps Script personal finance tracker that pulls transactions via Plaid, categorizes and ledgers them against an annual budget, and emails a daily deterministic + LLM-narrated summary. Treated as a separate project/repo from the Toolsmith Agent project, even though it currently lives alongside it — it only reuses Toolsmith Agent's `llm()` function.

See also: [ARCHITECTUREOVERVIEW.md](./ARCHITECTUREOVERVIEW.md) · [CONTRIBUTING.md](./CONTRIBUTING.md)

## What this is
A daily-cron finance tracker: it syncs bank transactions from Plaid, categorizes them against user-defined rules, rolls them up into annual/monthly budget sheets, and emails a summary. Unrecognized transactions go into an "Unknown" bucket that the user resolves by **replying to email**, which the system parses and learns from.

## Architecture philosophy
- **Deterministic core, LLM only at the edges.** There is no LLM tool-orchestration/agent loop here (unlike Toolsmith Agent). The LLM is used one-shot, for two things only:
  1. Writing the qualitative email summary/feedback.
  2. Parsing free-form email replies into structured category assignments.
- All percentages, rollups, and pacing math are computed in code — never left to the LLM.

## Notable design choices
- **401k** is pre-tax/invisible to Plaid and intentionally not tracked.
- **Brokerage/Roth IRA** transfers could be tracked but deliberately aren't — no budget tension to reconcile, just a known fixed transfer.
- **Category Rules** are AND-matched across three Plaid fields (Merchant Name Contains, Plaid Category primary, Plaid Category detailed) plus a Target Sheet/Field, so the same merchant can route differently depending on context (e.g. Target for groceries vs. office supplies), and loan-payment types (credit card vs. mortgage vs. student/personal loan) can be told apart.
- A rule's Target Sheet can be `"Ignore"` — matching transactions (e.g. credit card payments, already counted when the original purchase posted) are skipped entirely: not tracked, not ledgered, not sent to Unknown.
- Resolving an Unknown transaction auto-creates a matching Category Rule (if one doesn't already exist), so repeat merchants stop needing a reply.

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
| `GROQ_API_KEY` | Auth for the shared `llm()` (summary writing, reply parsing) | Yes |
| `FINANCE_SPREADSHEET_NAME` | Points the script at the active finance spreadsheet; repointed by `startFreshFinanceSheet()` | Yes |
| `SYNC_START_DATE` | Bounds how far back Plaid sync pulls history (e.g. Jan 1 of current year) | Yes |

> Note: property names above reflect the project's known architecture (`FINANCE_SPREADSHEET_NAME` and `SYNC_START_DATE` are confirmed from project notes; Plaid/Groq credential names are the conventional Plaid/Groq naming and should be confirmed against your live Script Properties).

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
7. **Run `setupFinanceSheets()`** (creates the Unknown sheet and any other structural setup) if starting fresh.
8. **Wire up triggers**: a daily trigger for `dailyFinanceCheck()`, and a 15-minute trigger for `checkForCategoryReplies()`.
9. **Populate initial Category Rules** for known recurring merchants, including any `"Ignore"` rules (e.g. credit card payment transactions).
10. **First run / backfill**: trigger a sync manually and watch for Unknown-bucket entries; reply to the summary email to resolve them and start building out Category Rules organically.

### Keeping the repo and live script in sync
Pick one direction as source of truth to avoid silently overwriting changes:
- **Repo is source of truth**: edit `.gs` files locally, `clasp push` to deploy.
- **Apps Script UI is source of truth**: edit in the browser editor, `clasp pull` before your next local edit, then commit.

### Starting a new year / fresh sheet
Use `startFreshFinanceSheet(newSpreadsheetName)` — it repoints `FINANCE_SPREADSHEET_NAME` at the new sheet and clears sync cursors (bounded by `SYNC_START_DATE`) while keeping the same linked Plaid bank connections.
