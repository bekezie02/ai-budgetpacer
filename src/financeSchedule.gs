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

// =========================================================================
// Daily finance sync trigger — pulls transactions, updates the Monthly
// LifeStyle/Mandatory sheets and fixed-pool ledgers, and only emails a
// summary if something actually changed. See dailyFinanceCheck() in
// financeEmail.gs.
// =========================================================================
function setDailyFinanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "dailyFinanceCheck") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("dailyFinanceCheck")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}

function resetDailyFinanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "dailyFinanceCheck") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  setDailyFinanceTrigger();
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

  if (threads.length === 0) return; // Nothing to do

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
  const threadId = properties.getProperty('FINANCE_REPLY_THREAD_ID');
  if (!threadId) return; // nothing outstanding to watch for

  const thread = GmailApp.getThreadById(threadId);
  if (!thread) {
    properties.deleteProperty('FINANCE_REPLY_THREAD_ID');
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
