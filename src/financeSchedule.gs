function setFinancialTrigger() {
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
  if (!event || !event.range || event.range.getSheet().getName() !== LEDGER_SHEET) return;
  const range = event.range;
  const firstColumn = range.getColumn();
  const lastColumn = firstColumn + range.getNumColumns() - 1;
  if (firstColumn > 9 || lastColumn < 9 || range.getLastRow() < 2) return;

  const sheet = range.getSheet();
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getLastRow();
  const isSingleCellEdit = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const previousValue = isSingleCellEdit ? event.oldValue : undefined;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    for (let row = firstRow; row <= lastRow; row++) {
      try {
        restoreTransactionLedgerOverrideSelection(row, row === firstRow ? previousValue : undefined);
      } catch (error) {
        const detail = error && error.message ? error.message : String(error);
        console.error(`Could not restore the override selection on ledger row ${row}: ${detail}`);
      }
    }
    sheet.getParent().toast(
      "A finance sync or replacement is running. The edit was not applied; please retry shortly.",
      "Transaction category not changed",
      6
    );
    return;
  }

  try {
    for (let row = firstRow; row <= lastRow; row++) {
      const selectedLabel = sheet.getRange(row, 9).getValue();
      try {
        processTransactionLedgerOverrideEdit(row, selectedLabel);
      } catch (error) {
        try {
          restoreTransactionLedgerOverrideSelection(row, row === firstRow ? previousValue : undefined);
        } catch (restoreError) {
          const restoreDetail = restoreError && restoreError.message
            ? restoreError.message
            : String(restoreError);
          console.error(`Could not restore the override selection on ledger row ${row}: ${restoreDetail}`);
        }
        const detail = error && error.message ? error.message : String(error);
        console.error(`Could not apply the transaction override on ledger row ${row}: ${detail}`);
        sheet.getParent().toast(
          `Row ${row}: ${detail}`,
          "Transaction category not changed",
          8
        );
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// Daily finance sync trigger — pulls transactions, updates the Monthly
// LifeStyle/Mandatory sheets and fixed-pool ledgers, and only emails a
// summary if something actually changed. See dailyFinanceCheck() in
// financeEmail.gs.
// =========================================================================
function setDailyFinanceTrigger() {
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

// =========================================================================
// Category-reply trigger — watches the single outstanding finance-summary
// thread (if any) for replies categorizing Unknown-bucket transactions.
// See checkForCategoryReplies() below.
// =========================================================================
function setCategoryReplyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkForCategoryReplies") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("checkForCategoryReplies")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function resetCategoryReplyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkForCategoryReplies") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  setCategoryReplyTrigger();
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
  checkPendingPlaidReplacementApproval();
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
  if (
    financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_ITEM_ID') === itemId &&
    financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_THREAD_ID')
  ) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const currentItemId = financeScriptProperties.getProperty('PLAID_PENDING_REPLACEMENT_ITEM_ID');
    if (!currentItemId || currentItemId !== itemId) return;
    if (
      financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_ITEM_ID') === itemId &&
      financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_THREAD_ID')
    ) return;

    const result = stagePendingPlaidItemTransactionsLocked();
    if (result.status === "in_progress") return;

    financeScriptProperties.deleteProperty('PLAID_REPLACEMENT_STAGING_LAST_ERROR');
    sendPlaidReplacementReviewEmail(itemId, linkMessage);
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

function sendPlaidReplacementReviewEmail(itemId, linkMessage) {
  const subject = `Plaid Replacement Review: ${itemId}`;
  const reviewSheet = getPlaidReplacementReviewSheet(itemId);
  const reviewUrl = `${getFinanceSpreadsheet().getUrl()}#gid=${reviewSheet.getSheetId()}`;
  const body = `${linkMessage || `Transaction staging is complete for Plaid Item ${itemId}.`}\n\n` +
    `Review the staged transactions here: ${reviewUrl}\n` +
    `All available transactions for the current-year window have been staged. Choose the old bank/Item to replace from the ordered dropdown in cell B1; each option shows its institution name and Item ID. Decision is pre-filled: a unique exact date, amount, and description match is REPLACE <old transaction ID>; all other transactions default to IMPORT. Change IMPORT to SKIP in column A for any transaction you do not want to import. ` +
    `Status REVIEW flags rows with multiple possible matches in column K; check these carefully because their default IMPORT could create a duplicate. REPLACE removes the old Ledger/Unknown entry before applying the new Item's transaction ID. ` +
    `Then reply to this email with exactly:\nFINALIZE ${itemId}\n\n` +
    `The replacement will not be finalized until every staged transaction has a valid decision.\n\n` +
    `To cancel and start over, reply with exactly:\nCANCEL ${itemId}\n` +
    `Cancellation deletes this replacement's review sheet and pending local Item credentials. It is refused if any transaction has already been imported or replaced, and it does not remove the Item from Plaid.`;

  GmailApp.sendEmail(MY_EMAIL, subject, body);
  const sentThreads = GmailApp.search(
    `in:sent to:${MY_EMAIL} subject:"${subject}" newer_than:1d`,
    0,
    10
  );
  const sentThread = sentThreads.find(thread => thread.getFirstMessageSubject() === subject);
  if (!sentThread) {
    throw new Error(`Sent the Plaid replacement review email, but could not locate its Gmail thread for approval tracking (${itemId}).`);
  }

  sentThread.markRead();
  financeScriptProperties.setProperty('PLAID_REPLACEMENT_APPROVAL_THREAD_ID', sentThread.getId());
  financeScriptProperties.setProperty('PLAID_REPLACEMENT_APPROVAL_ITEM_ID', itemId);
}

function checkPendingPlaidReplacementApproval() {
  const threadId = financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_THREAD_ID');
  const expectedItemId = financeScriptProperties.getProperty('PLAID_REPLACEMENT_APPROVAL_ITEM_ID');
  const pendingItemId = financeScriptProperties.getProperty('PLAID_PENDING_REPLACEMENT_ITEM_ID');
  if (!threadId || !expectedItemId || expectedItemId !== pendingItemId) return;

  const thread = GmailApp.getThreadById(threadId);
  if (!thread) {
    throw new Error(`Could not find the Plaid replacement approval thread for Item ${expectedItemId}.`);
  }

  const reply = thread.getMessages().find(message => {
    if (!message.isUnread()) return false;
    const from = message.getFrom();
    const emailMatch = from.match(/<([^>]+)>/);
    const senderEmail = (emailMatch ? emailMatch[1] : from).trim().toLowerCase();
    return senderEmail === MY_EMAIL.toLowerCase();
  });
  if (!reply) {
    if (financeScriptProperties.getProperty('PLAID_REPLACEMENT_FINALIZATION_ITEM_ID') === expectedItemId) {
      processPendingPlaidReplacementFinalization(thread, expectedItemId);
    }
    return;
  }

  reply.markRead();
  const replyCommand = getPlaidReplacementReplyCommand(reply);
  if (replyCommand.toUpperCase() === `CANCEL ${expectedItemId}`.toUpperCase()) {
    let result;
    try {
      result = cancelPlaidItemReplacement();
    } catch (error) {
      thread.reply(
        `The Plaid Item replacement was not cancelled:\n${error.message || error}\n\n` +
        `If no transactions have been imported or replaced, reply again with exactly:\nCANCEL ${expectedItemId}`
      );
      return;
    }

    thread.reply(
      `Plaid Item replacement cancelled.\n` +
      `Pending Item: ${expectedItemId}\n` +
      `Review sheet deleted: ${result.review_sheet_deleted ? "yes" : "no"}\n` +
      `The Item itself was not removed from Plaid.`
    );
    return;
  }

  if (replyCommand.toUpperCase() !== `FINALIZE ${expectedItemId}`.toUpperCase()) {
    thread.reply(
      `No action was taken. Put exactly one of these commands on the first nonblank line of your reply:\n` +
      `FINALIZE ${expectedItemId}\n` +
      `CANCEL ${expectedItemId}`
    );
    return;
  }

  financeScriptProperties.setProperty('PLAID_REPLACEMENT_FINALIZATION_ITEM_ID', expectedItemId);
  processPendingPlaidReplacementFinalization(thread, expectedItemId);
}

function processPendingPlaidReplacementFinalization(thread, itemId) {
  try {
    const result = finalizePlaidItemReplacement();
    if (result.status === "in_progress") {
      if (financeScriptProperties.getProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID") !== itemId) {
        thread.reply(
          `Plaid Item replacement finalization is continuing in resumable batches.\n` +
          `Item: ${itemId}\n` +
          `No further reply is needed; the scheduled trigger will continue and email when complete.`
        );
        financeScriptProperties.setProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID", itemId);
      }
      return;
    }

    const sheetCleanup = result.review_sheet_deleted
      ? "The replacement review sheet was deleted."
      : `The replacement is complete, but its review sheet could not be deleted: ${result.review_sheet_delete_error}`;
    thread.reply(
      `Plaid Item replacement completed.\n` +
      `Old Item: ${result.old_item_id}\n` +
      `New Item: ${result.new_item_id}\n` +
      `The new Item's access token and staged cursor are retained; the old Item's token and cursor were removed.\n` +
      sheetCleanup
    );
  } catch (error) {
    financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID");
    financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID");
    thread.reply(`The Plaid Item replacement was not finalized:\n${error.message || error}\n\nAfter addressing the issue, reply again with exactly:\nFINALIZE ${itemId}`);
  }
}

function getPlaidReplacementReplyCommand(messageObject) {
  const replyText = extractNewReplyOnly(messageObject);
  const firstReplyLine = replyText.split(/\r?\n/).find(line => line.trim() && !line.trim().startsWith(">"));
  return firstReplyLine ? firstReplyLine.trim() : "";
}

function extractNewReplyOnly(messageObject) {
  // 1. Use plain text body to immediately save ~80% of token weight
  let body = messageObject.getPlainBody();
  
  // 2. Common headers that mark where the old thread history begins
  const quoteMarkers = [
    /\bOn\s+.*\s+wrote:/i,                                // Gmail standard: "On [Date], [Name] wrote:"
    /-----Original Message-----/i,                        // Outlook standard
    /From:\s*.*[\r\n]+Sent:\s*/i,                         // Windows Live / Outlook variant
    /From:\s*.*\s*<.*>\s*[\r\n]+Date:/i,                  // Apple Mail variant
    /_\s*[\r\n]+From:\s*/i                                // General webmail delimiters
  ];
  
  // 3. Split the text at the first matching marker found
  for (let marker of quoteMarkers) {
    let parts = body.split(marker);
    if (parts.length > 1) {
      body = parts[0]; // Keep only the text BEFORE the marker
      break; 
    }
  }
  return body.trim();
}

// =========================================================================
// Watches the single stored finance-summary thread ID for unread replies,
// parses them into {transactionId, category} via the LLM, resolves any
// that match the exact-match category whitelist, emails a per-line
// confirmation, and — if anything was resolved — sends a refreshed summary
// email right away rather than waiting for tomorrow's cron.
// =========================================================================
function checkForCategoryReplies() {
  const threadId = financeScriptProperties.getProperty('FINANCE_REPLY_THREAD_ID');
  if (!threadId) return; // nothing outstanding to watch for

  const thread = GmailApp.getThreadById(threadId);
  if (!thread) {
    financeScriptProperties.deleteProperty('FINANCE_REPLY_THREAD_ID');
    return;
  }

  const unreadReplies = thread.getMessages().filter(m => m.isUnread());
  if (unreadReplies.length === 0) return;
  const labelMap = buildCategoryLabelMap(listValidCategoryTargets());
  const validLabels = Object.keys(labelMap);
  const results = [];
  let resolvedCount = 0;
  for(let msg of unreadReplies) {
    // NEW LINE: Extracts text-only and slices off historical thread junk
    // 1. Clean the text using plain text and regex splitting
    const cleanReplyText = extractNewReplyOnly(msg);
    // 2. SAFETY GUARD: 1 token is roughly 4 characters. 
    // 8,000 token limit means a max of ~32,000 characters. 
    // Since an automated categorical reply should be very short, let's cap it strictly at 4,000 characters (~1,000 tokens).
    if (cleanReplyText.length > 4000) {
      console.warn(`Skipping message ${msg.getId()} — Body size (${cleanReplyText.length} chars) is still too large for LLM.`);
      msg.markRead(); // Mark read so it doesn't loop forever, or remove this line to keep it unread.
      continue;       // Safely skip this iteration and check the next email
    }
    const parsed = parseCategoryReply(cleanReplyText, validLabels);
    msg.markRead();
    if (!parsed.success) {
      console.warn("Failed to parse category reply:", parsed.error);
      continue;
    }
    parsed.data.forEach(entry => {
      console.log(entry)
      const target = labelMap[String(entry.category || "").trim().toLowerCase()];
      if (!target) {
        results.push(`${entry.transactionId}: "${entry.category}" isn't a valid category — left uncategorized.`);
        return;
      }
      const outcome = resolveUnknownTransaction(entry.transactionId, target.targetSheet, target.targetField);
      if (outcome.success) resolvedCount++;
      results.push(outcome.success
        ? `${entry.transactionId}: assigned to ${entry.category}.`
        : `${entry.transactionId}: failed — ${outcome.error}`);

      if (outcome.success && outcome.alsoResolved && outcome.alsoResolved.length > 0) {
        outcome.alsoResolved.forEach(match => {
          resolvedCount++;
          results.push(`${match.transaction_id}: also assigned to ${entry.category} (same merchant/category as ${entry.transactionId}).`);
        });
      }
    });
  };

  if (results.length > 0) {
    MailApp.sendEmail({ to: MY_EMAIL, subject: "Finance categorization update", body: results.join("\n") });
  }

  if (resolvedCount > 0) {
    sendFinanceSummaryEmail();
  }
}