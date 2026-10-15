# Project Changes Guide

This guide explains the main changes made to AI BudgetPacer during this work. It is written for someone who is new to the source code.

## In a nutshell

The project still gets transactions from Plaid, categorizes them, and updates the budget spreadsheet. The changes make it easier to:

- fetch the current year's history when connecting a bank;
- finish a large transaction sync without repeatedly running it by hand;
- safely review and replace a connected Plaid Item; and
- move one transaction to a different budget without changing the rule for future transactions.

## 1. Fetch transactions from the start of the year

### What changed

When a new Plaid Link connection is created, the script calculates how many days have elapsed since January 1 of the current year and requests that much transaction history. The start date is calculated from the current year rather than being a fixed month or number of days.

The normal sync also applies transactions from January 1 onward to the budget spreadsheet. Plaid's saved cursor records how far the script has read for each connected Item. `resyncTransactionsFromStart()` can clear those cursors and replay the history Plaid makes available to the Item.

### Why

A fixed history request, such as 243 days, can miss transactions from January. Clearing a cursor alone cannot recover history that Plaid did not provide when the Item was first connected. Calculating the request from January 1 helps new connections include the full year-to-date period.

### How it works

- `runStep1()` creates the regular Link token and calculates `days_requested`.
- `runDailyFinanceSync()` fetches pages from Plaid and applies transactions from the configured start date.
- `resyncTransactionsFromStart()` clears saved cursors but leaves existing spreadsheet data in place; the normal transaction logic updates existing rows rather than intentionally clearing the budget.

Relevant source: [financeUtils.gs](./src/financeUtils.gs) and [financeCore.gs](./src/financeCore.gs).

## 2. Continue large daily syncs automatically

### What changed

The daily sync now saves each page's cursor and stops at an internal five-minute work limit. A one-minute continuation trigger resumes the sync while it is still catching up. Once all active Items are current, the script sends a completion email; it does not send a partial finance summary for every continuation.

Errors such as a missing Item access token or a Plaid sync failure are logged and reported by email, and the continuation trigger retries.

### Why

An initial sync can contain many transactions and exceed one Apps Script execution. Saving progress between pages allows the next execution to continue instead of requiring the user to run a sync function again and again.

### How to start it

1. Deploy the updated Apps Script code.
2. Run `setDailyFinanceTrigger()` once to install the daily trigger and the one-minute continuation trigger.
3. Run `dailyFinanceCheck()` once to start the initial sync. The continuation trigger handles remaining pages.

The five-minute limit leaves room before Apps Script's execution timeout. If Apps Script stops an execution unexpectedly, the last saved cursor is retained. If the sync does not resume, run `dailyFinanceCheck()` again.

Relevant source: [financeCore.gs](./src/financeCore.gs), [financeEmail.gs](./src/financeEmail.gs), and [financeSchedule.gs](./src/financeSchedule.gs).

## 3. Review Plaid Item replacements before switching banks

### What changed

A replacement connection is staged in a separate `Plaid Review <Item ID>` sheet instead of being made active immediately. Its transactions are checked against existing transactions:

- A unique exact date, amount, and description match is suggested as `REPLACE`.
- A transaction with no exact match is suggested as `IMPORT`.
- Multiple possible matches are flagged for review.
- `SKIP` leaves a staged transaction out of the budget.

The review sheet's B1 dropdown lists existing banks and Item IDs in the same order as `PLAID_ITEM_IDS`. Bank names are looked up from Plaid and cached in a `PLAID_INSTITUTION_NAME_<Item ID>` Script Property. This value is an optional bank-name cache, not an access token or a required Item setting.

The review sheet now provides a **Replacement action** dropdown in cell B5. Choose **Finalize replacement** after reviewing all staged decisions or **Cancel replacement** to abandon the staged connection. Finalization is time-bounded and continues automatically on later trigger runs. An email is sent when staging starts, when review is ready, and when finalization completes.

### Why

Replacing a bank connection creates a new Plaid Item and new transaction IDs. Reviewing the staged data first helps avoid counting the same real-world purchase twice and makes it clear which existing bank connection will be replaced.

### How to use it

1. Send the replacement request email with subject `Plaid Item Replacement`.
2. Open the emailed Plaid Link and connect the bank.
3. In the generated review sheet, choose the existing bank/Item in B1.
4. Review transaction decisions; correct any suggestions, including changing a row to `SKIP` when appropriate.
5. Use the **Replacement action** dropdown in B5 to finalize or cancel. If finalization takes longer than one run, it resumes automatically.

Relevant source: [financeCore.gs](./src/financeCore.gs), [financeUtils.gs](./src/financeUtils.gs), and [financeSchedule.gs](./src/financeSchedule.gs).

## 4. Override the budget for one transaction

### What changed

The Transaction Ledger has an **Override Category** dropdown for each transaction. Selecting a different monthly category, fixed pool (such as Travel), or Ignore moves that transaction's amount from its current target to the selected target. It does not change the reusable Category Rule, so future dining transactions can still go to the usual eating-out budget.

The override is kept when Plaid later updates the transaction. Selecting **Use original category** clears it and restores the transaction's original target. The added default-target columns in the ledger store that original target.

### Why

Sometimes the regular merchant/category rule is correct most of the time, but one transaction needs different treatment—for example, a meal bought while traveling should count toward Travel. A per-transaction override handles that case without changing future categorization.

### How to turn it on

Deploy the updated code and run `setupFinanceSheets()` once. It updates an existing Transaction Ledger in place, adds the override columns and dropdown, and installs the spreadsheet edit trigger. Then select a value in that transaction's **Override Category** cell; no separate function call is needed for each change.

Relevant source: [financeCore.gs](./src/financeCore.gs) and [financeSchedule.gs](./src/financeSchedule.gs).

## 5. Assign unassigned transactions from the sheet

### What changed

The finance email still lists unassigned transactions, but now links directly to the **Uncategorized Transactions** sheet. Each transaction has an **Assign to Category** dropdown. Choosing a category applies it to that transaction and other unassigned transactions with the same merchant and Plaid primary and detailed categories. The script also creates a matching Category Rule for future transactions.

The sheet displays a confirmation after assignment; selecting a dropdown value does not send a follow-up email. Email replies are not processed as category assignments.

### Why

Users can assign transactions where they already review them, without copying transaction IDs into an email. Matching transactions are still handled together, preserving the existing categorization behavior.

### How it works

The spreadsheet edit trigger calls the existing `resolveUnknownTransaction()` logic. That logic applies the target, removes resolved rows from Unknown, creates a reusable Category Rule, and resolves identical merchant/category matches. `setupFinanceSheets()` adds the dropdown column to an existing sheet and installs the trigger.

Relevant source: [financeCore.gs](./src/financeCore.gs), [financeEmail.gs](./src/financeEmail.gs), and [financeSchedule.gs](./src/financeSchedule.gs).

## 6. Automatically add spreadsheet rows when needed

### What changed

The script checks whether a sheet has enough rows before it adds transactions or monthly sections. If it is about to run out of rows, it adds more automatically instead of stopping and requiring someone to click **Add more rows**.

This supports automatically populated sheets such as monthly budget sections, fixed-pool ledgers, the Transaction Ledger, Uncategorized Transactions, Category Rules, and Plaid replacement review sheets.

### Why

Google Sheets has a finite number of rows allocated to each sheet. Apps Script cannot write past that limit unless it first expands the sheet, so a growing transaction history could otherwise stop an automated sync.

### How it works

`ensureSheetHasRows()` compares the required row with the sheet's current capacity and inserts additional rows when needed. Shared append operations and sheet-specific workflows call this helper before writing. Users should not need to expand rows manually for these automated inserts.

Relevant source: [financeCore.gs](./ai-budgetpacer/src/financeCore.gs).

## Notes for maintainers

- The `.gs` files share a global scope in Apps Script. The shared Script Properties object is named `financeScriptProperties` to avoid collisions with other top-level declarations.
- Script Properties such as access tokens and cursors belong to the live Apps Script project; they should not be added to source files.
- The code and documentation changes have been checked locally for JavaScript syntax and whitespace errors. They have not been deployed or tested against the live Plaid account, Apps Script triggers, or finance spreadsheet.
