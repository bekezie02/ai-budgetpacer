function setFinancialTrigger() {
  removeCategoryReplyTriggers();
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkInboxForTriggers") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("checkInboxForTriggers")
    .timeBased()
    .everyMinutes(1)
    .create()
}

function resetFinancialTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkInboxForTriggers") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  setFinancialTrigger();
}

function ensureTransactionLedgerEditTrigger(spreadsheet) {
  const spreadsheetId = spreadsheet.getId();
  const triggerExists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === "handleTransactionLedgerEdit" &&
    trigger.getTriggerSourceId() === spreadsheetId
  );
  if (!triggerExists) {
    ScriptApp.newTrigger("handleTransactionLedgerEdit")
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();
  }
}

function handleTransactionLedgerEdit(event) {
  if (!event || !event.range) return;
  const range = event.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  if (sheetName.startsWith(PLAID_REPLACEMENT_REVIEW_PREFIX)) {
    handlePlaidReplacementActionEdit(event);
    return;
  }
  if (sheetName !== LEDGER_SHEET && sheetName !== UNKNOWN_SHEET) return;
  const categoryColumn = sheetName === LEDGER_SHEET
    ? LEDGER_OVERRIDE_COLUMN
    : UNKNOWN_CATEGORY_COLUMN;
  const firstColumn = range.getColumn();
  const lastColumn = firstColumn + range.getNumColumns() - 1;
  if (firstColumn > categoryColumn || lastColumn < categoryColumn || range.getLastRow() < 2) return;

  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();
  const isSingleCellEdit = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const previousValue = isSingleCellEdit ? event.oldValue : undefined;
  const edits = [];
  for (let row = firstRow; row <= lastRow; row++) {
    const values = sheet.getRange(row, 1, 1, categoryColumn).getValues()[0];
    edits.push({
      row,
      transactionId: values[0],
      selectedLabel: values[categoryColumn - 1]
    });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    edits.forEach(edit => restoreFinanceCategorySelection(
      sheet,
      sheetName,
      edit,
      isSingleCellEdit ? previousValue : undefined
    ));
    sheet.getParent().toast(
      "A finance sync or replacement is running. The category change was not applied; please retry shortly.",
      "Transaction not categorized",
      6
    );
    return;
  }

  try {
    let assignedCount = 0;
    const assignedByCategory = {};
    for (const edit of edits) {
      if (!edit.transactionId || !edit.selectedLabel) continue;
      try {
        if (sheetName === LEDGER_SHEET) {
          processTransactionLedgerOverrideEdit(edit.row, edit.selectedLabel);
        } else {
          if (!getUnknownTransaction(edit.transactionId)) continue;
          const outcome = assignUnknownTransactionFromSheet(edit.transactionId, edit.selectedLabel);
          const count = 1 + outcome.alsoResolved.length;
          assignedCount += count;
          assignedByCategory[edit.selectedLabel] = (assignedByCategory[edit.selectedLabel] || 0) + count;
        }
      } catch (error) {
        const detail = error && error.message ? error.message : String(error);
        console.error(`Could not apply the category selection on ${sheetName} row ${edit.row}: ${detail}`);
        restoreFinanceCategorySelection(
          sheet,
          sheetName,
          edit,
          isSingleCellEdit ? previousValue : undefined
        );
        sheet.getParent().toast(
          `Row ${edit.row}: ${detail}`,
          "Transaction not categorized",
          8
        );
      }
    }
    if (sheetName === UNKNOWN_SHEET && assignedCount > 0) {
      const assignmentSummary = Object.entries(assignedByCategory)
        .map(([category, count]) => `${count} to ${category}`)
        .join("; ");
      sheet.getParent().toast(
        `Assigned ${assignedCount} transaction${assignedCount === 1 ? "" : "s"} (${assignmentSummary}).`,
        "Unassigned transactions updated",
        6
      );
    }
  } finally {
    lock.releaseLock();
  }
}

function handlePlaidReplacementActionEdit(event) {
  const range = event.range;
  if (
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1 ||
    range.getRow() !== PLAID_REPLACEMENT_ACTION_ROW ||
    range.getColumn() !== PLAID_REPLACEMENT_ACTION_COLUMN
  ) return;

  const sheet = range.getSheet();
  const action = String(range.getValue() || "").trim();
  const itemId = String(sheet.getRange(2, 2).getValue() || "").trim();
  if (!itemId || !action) return;
  range.clearContent();

  try {
    if (action === PLAID_REPLACEMENT_FINALIZE_ACTION) {
      const result = finalizePlaidItemReplacement();
      if (result.status === "in_progress") {
        sheet.getParent().toast(
          "Finalization is running in batches and will continue automatically.",
          "Plaid replacement in progress",
          8
        );
        return;
      }
      sendPlaidReplacementFinalizationEmail(result);
      sheet.getParent().toast("Plaid Item replacement completed.", "Replacement complete", 8);
      return;
    }

    if (action === PLAID_REPLACEMENT_CANCEL_ACTION) {
      cancelPlaidItemReplacement();
      sheet.getParent().toast(
        "The staged replacement was cancelled. The active Item and finance data were not changed.",
        "Replacement cancelled",
        8
      );
      return;
    }
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    const currentSheet = getFinanceSpreadsheet().getSheetByName(sheet.getName());
    if (currentSheet) currentSheet.getRange(4, 2).setValue(`Last action failed: ${detail}`);
    sheet.getParent().toast(detail, "Plaid replacement action failed", 10);
    console.error(`Could not apply Plaid replacement action for Item ${itemId}: ${detail}`);
  }
}

function restoreFinanceCategorySelection(sheet, sheetName, edit, previousValue) {
  try {
    if (sheetName === LEDGER_SHEET) {
      restoreTransactionLedgerOverrideSelection(
        edit.row,
        previousValue === undefined ? undefined : previousValue
      );
      return;
    }
    const unknowns = getAllUnknownTransactions();
    const currentRow = unknowns.findIndex(txn => txn.transaction_id === edit.transactionId) + 2;
    if (currentRow >= 2) {
      sheet.getRange(currentRow, UNKNOWN_CATEGORY_COLUMN)
        .setValue(previousValue === undefined ? "" : previousValue);
    }
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    console.error(`Could not restore the category selection on ${sheetName} row ${edit.row}: ${detail}`);
  }
}

// =========================================================================
// Daily finance sync trigger — pulls transactions, updates the Monthly
// LifeStyle/Mandatory sheets and fixed-pool ledgers, and only emails a
// summary if something actually changed. See dailyFinanceCheck() in
// financeEmail.gs.
// =========================================================================
function setDailyFinanceTrigger() {
  removeCategoryReplyTriggers();
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (
      trigger.getHandlerFunction() === "dailyFinanceCheck" ||
      trigger.getHandlerFunction() === "continueDailyFinanceSync"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("dailyFinanceCheck")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  ScriptApp.newTrigger("continueDailyFinanceSync")
    .timeBased()
    .everyMinutes(1)
    .create();
}

function resetDailyFinanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (
      trigger.getHandlerFunction() === "dailyFinanceCheck" ||
      trigger.getHandlerFunction() === "continueDailyFinanceSync"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  setDailyFinanceTrigger();
}

function continueDailyFinanceSync() {
  if (financeScriptProperties.getProperty("DAILY_FINANCE_BACKFILL_ACTIVE") !== "true") return;
  dailyFinanceCheck(true);
}

function removeCategoryReplyTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkForCategoryReplies") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  financeScriptProperties.deleteProperty("FINANCE_REPLY_THREAD_ID");
}

// =========================================================================
// TRIGGER 1: Run this on a 1-minute time-trigger to look for your trigger email
// =========================================================================
function checkInboxForTriggers() {
  // Look for unread emails from you with the subject "Plaid Link"
  const threads = GmailApp.search(`is:unread from:${MY_EMAIL} subject:"Plaid Link"`);

  for (const thread of threads) {
    // 1. Mark it read so it doesn't trigger again
    thread.markRead();

    // 2. Run Step 1 to generate the link
    const hostedLinkUrl = runStep1();

    // 3. Email the link back to yourself
    if (hostedLinkUrl) {
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "🔗 Action Required: Complete Your Plaid Authentication",
        htmlBody: `
          <p>You requested a new bank connection setup.</p>
          <p><strong>Step 1:</strong> Click the button below to sign into your bank:</p>
          <p><a href="${hostedLinkUrl}" style="background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">Log Into Bank</a></p>
          <p><em>Note: You have exactly 5 minutes to finish logging in before the script checks for completion.</em></p>
        `
      });
    }
    // PAUSE EXECUTION: Give yourself 5 minutes to tap the link and sign in
    Logger.log("Link emailed. Pausing script execution for 5 minutes...");
    Utilities.sleep(300000); // 300000 ms = 5 minutes
    // Call Step 2 automatically in the same container session!
    const outcome = runStep2();

    // Send final notification email based on results
    if (outcome.status === "success") {
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "✅ Plaid Success: Bank Account Linked Successfully",
        htmlBody: `<p>Success! The single-thread handshake completed cleanly. Your token is saved to script properties.</p>`
      });
    } else {
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "❌ Plaid Failure: Connection Setup Timed Out",
        htmlBody: `<p>The handshake failed. Reason given:</p><pre>${outcome.message}</pre><p>Please send a new "Plaid Link" email to restart.</p>`
      });
    }
  }

  // Replacement requests use a distinct subject so they cannot accidentally
  // take the normal path that immediately activates every new Plaid Item.
  const replacementThreads = GmailApp.search(
    `is:unread from:${MY_EMAIL} subject:"Plaid Item Replacement"`
  );

  for (const thread of replacementThreads) {
    thread.markRead();

    try {
      const hostedLinkUrl = startPlaidItemReplacementLink();
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "🔗 Action Required: Review New Plaid Item Connection",
        htmlBody: `
          <p>You requested a new Plaid Item for transaction-history replacement.</p>
          <p><strong>Step 1:</strong> Complete the bank connection using the link below:</p>
          <p><a href="${hostedLinkUrl}" style="background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">Connect Bank</a></p>
          <p>The script will check for completion automatically for up to 25 minutes. The new Item will be staged for review and will not replace the active Item automatically.</p>
        `
      });
      Logger.log("Replacement Link emailed. The next inbox-trigger runs will check for completion.");
    } catch (error) {
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "❌ Could Not Start Plaid Item Replacement",
        htmlBody: `<p>The replacement request could not be started:</p><pre>${error.message || error}</pre>`
      });
    }
  }

  checkPendingPlaidReplacementLink();
  checkPendingPlaidReplacementStaging();
  continuePendingPlaidReplacementFinalization();
}

function checkPendingPlaidReplacementLink() {
  if (financeScriptProperties.getProperty('PLAID_LINK_MODE') !== 'REPLACEMENT_REVIEW') return;

  const outcome = runStep2();
  if (outcome.status === "pending") return;

  if (outcome.status === "success") {
    financeScriptProperties.deleteProperty('PLAID_REPLACEMENT_LINK_LAST_ERROR');
    const itemId = financeScriptProperties.getProperty('PLAID_PENDING_REPLACEMENT_ITEM_ID');
    if (!itemId) throw new Error("Plaid reported replacement success, but the pending Item ID is missing.");
    Logger.log(`Plaid Item ${itemId} linked; transaction staging will run automatically.`);
    return;
  }

  if (!financeScriptProperties.getProperty('PLAID_REPLACEMENT_LINK_LAST_ERROR')) {
    MailApp.sendEmail({
      to: MY_EMAIL,
      subject: "❌ Plaid Item Replacement Link Needs Attention",
      htmlBody: `<p>The replacement Link check failed:</p><pre>${outcome.message}</pre><p>The script will keep checking until the Link session times out.</p>`
    });
    financeScriptProperties.setProperty('PLAID_REPLACEMENT_LINK_LAST_ERROR', outcome.message || "Unknown Link error");
  }
}

function checkPendingPlaidReplacementStaging(linkMessage) {
  const itemId = financeScriptProperties.getProperty('PLAID_PENDING_REPLACEMENT_ITEM_ID');
  if (!itemId) return;
  if (financeScriptProperties.getProperty("PLAID_REPLACEMENT_STAGING_COMPLETE_ITEM_ID") === itemId) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const currentItemId = financeScriptProperties.getProperty('PLAID_PENDING_REPLACEMENT_ITEM_ID');
    if (!currentItemId || currentItemId !== itemId) return;
    if (financeScriptProperties.getProperty("PLAID_REPLACEMENT_STAGING_COMPLETE_ITEM_ID") === itemId) return;

    sendPlaidReplacementStagingStartedEmail(itemId);
    const result = stagePendingPlaidItemTransactionsLocked();
    if (result.status === "in_progress") return;

    financeScriptProperties.deleteProperty('PLAID_REPLACEMENT_STAGING_LAST_ERROR');
    sendPlaidReplacementReviewEmail(itemId, linkMessage);
    financeScriptProperties.setProperty("PLAID_REPLACEMENT_STAGING_COMPLETE_ITEM_ID", itemId);
  } catch (error) {
    const message = error.message || String(error);
    if (financeScriptProperties.getProperty('PLAID_REPLACEMENT_STAGING_LAST_ERROR') !== message) {
      MailApp.sendEmail({
        to: MY_EMAIL,
        subject: "❌ Plaid Item Replacement Staging Needs Attention",
        body:
          `Automatic staging could not continue for Plaid Item ${itemId}:\n${message}\n\n` +
          `The scheduled trigger will retry. You can also run stagePendingPlaidItemTransactions() manually.`
      });
      financeScriptProperties.setProperty('PLAID_REPLACEMENT_STAGING_LAST_ERROR', message);
    }
  } finally {
    lock.releaseLock();
  }
}

function sendPlaidReplacementStagingStartedEmail(itemId) {
  if (financeScriptProperties.getProperty("PLAID_REPLACEMENT_STAGING_STARTED_NOTIFIED_ITEM_ID") === itemId) return;
  const reviewSheet = getPlaidReplacementReviewSheet(itemId);
  reviewSheet.getRange(3, 2).setValue("Staging in progress — fetching available transactions.");
  const reviewUrl = `${getFinanceSpreadsheet().getUrl()}#gid=${reviewSheet.getSheetId()}`;
  MailApp.sendEmail({
    to: MY_EMAIL,
    subject: `Plaid Replacement Staging Started: ${itemId}`,
    body:
      `Transaction staging has started for the new Plaid Item ${itemId}.\n\n` +
      `The script is fetching and adding transactions to the review sheet in resumable batches. ` +
      `You will receive another email when staging is complete.\n\n` +
      `Review sheet: ${reviewUrl}`
  });
  financeScriptProperties.setProperty("PLAID_REPLACEMENT_STAGING_STARTED_NOTIFIED_ITEM_ID", itemId);
}

function sendPlaidReplacementReviewEmail(itemId, linkMessage) {
  const subject = `Plaid Replacement Review: ${itemId}`;
  const reviewSheet = getPlaidReplacementReviewSheet(itemId);
  const reviewUrl = `${getFinanceSpreadsheet().getUrl()}#gid=${reviewSheet.getSheetId()}`;
  const body = `${linkMessage || `Transaction staging is complete for Plaid Item ${itemId}.`}\n\n` +
    `Review the staged transactions here: ${reviewUrl}\n` +
    `All available transactions for the current-year window have been staged. Choose the old bank/Item to replace from the ordered dropdown in cell B1; each option shows its institution name and Item ID. Decision is pre-filled: a unique exact date, amount, and description match is REPLACE <old transaction ID>; all other transactions default to IMPORT. Change IMPORT to SKIP in column A for any transaction you do not want to import. ` +
    `Status REVIEW flags rows with multiple possible matches in column K; check these carefully because their default IMPORT could create a duplicate. REPLACE removes the old Ledger/Unknown entry before applying the new Item's transaction ID. ` +
    `When ready, use the Replacement action dropdown in cell B5 to finalize or cancel. Finalization requires every staged transaction to have a valid decision and resumes automatically if it takes multiple batches. ` +
    `Cancellation deletes the review sheet and pending local Item credentials, is refused after any transaction has been imported or replaced, and does not remove the Item from Plaid.`;

  MailApp.sendEmail(MY_EMAIL, subject, body);
}

function continuePendingPlaidReplacementFinalization() {
  const itemId = financeScriptProperties.getProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID");
  if (!itemId) return;
  try {
    const result = finalizePlaidItemReplacement();
    if (result.status === "in_progress") return;
    sendPlaidReplacementFinalizationEmail(result);
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    const reviewSheet = getFinanceSpreadsheet().getSheetByName(getPlaidReplacementReviewSheetName(itemId));
    if (reviewSheet) reviewSheet.getRange(4, 2).setValue(`Finalization paused: ${detail}`);
    console.error(`Plaid replacement finalization paused for Item ${itemId}: ${detail}`);
    MailApp.sendEmail({
      to: MY_EMAIL,
      subject: `Plaid Replacement Finalization Needs Attention: ${itemId}`,
      body:
        `Finalization paused for Plaid Item ${itemId}:\n${detail}\n\n` +
        `After correcting the issue, choose "Finalize replacement" again from cell B5.`
    });
  }
}

function sendPlaidReplacementFinalizationEmail(result) {
  const sheetCleanup = result.review_sheet_deleted
    ? "The replacement review sheet was deleted."
    : `The replacement is complete, but its review sheet could not be deleted: ${result.review_sheet_delete_error}`;
  try {
    MailApp.sendEmail({
      to: MY_EMAIL,
      subject: `Plaid Item Replacement Completed: ${result.new_item_id}`,
      body:
        `Plaid Item replacement completed.\n` +
        `Old Item: ${result.old_item_id}\n` +
        `New Item: ${result.new_item_id}\n` +
        `The new Item's access token and staged cursor are retained; the old Item's token and cursor were removed.\n` +
        sheetCleanup
    });
  } catch (error) {
    const detail = error && error.message ? error.message : String(error);
    console.error(`Replacement completed, but its notification email could not be sent: ${detail}`);
  }
}