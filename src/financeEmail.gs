// =========================================================================
// FINANCE EMAIL — uses the LLM (via llm() from util.gs, same Apps Script
// project) in exactly two places: generateFinanceEmailSummary() and
// parseCategoryReply(). Everything else here is deterministic — the model
// is only asked to write feedback about numbers it's given, or to parse a
// reply against a fixed whitelist, never to decide what the numbers are
// or which category is "closest."
// =========================================================================
const ANNUAL_SAVINGS_TARGET_LABEL = "Projected Annual Savings";
const MANDATORY_ANNUAL_TARGET_LABEL = "Total Mandatory Costs Excluding Fixed Amounts";
const LIFESTYLE_ANNUAL_ESTIMATE_ANCHOR = "variable / lifestyle costs"; // substring match on the block header
 
function getLabeledValue(sheet, label) {
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim().toLowerCase() === label.toLowerCase()) {
      return Number(sheet.getRange(i + 1, 2).getValue()) || 0;
    }
  }
  return null;
}
 
function getAnnualSavingsTarget() {
  return getLabeledValue(getFinanceSheet("Annual Overview"), ANNUAL_SAVINGS_TARGET_LABEL);
}
 
function getMandatoryAnnualTarget() {
  return getLabeledValue(getFinanceSheet("Annual Overview"), MANDATORY_ANNUAL_TARGET_LABEL);
}
 
/**
 * The Lifestyle annual estimate (VARIABLE / LIFESTYLE COSTS (ANNUAL ESTIMATE))
 * lives in the Annual Overview sheet, ending in a "Total Variable Costs"
 * line. Anchoring on the block header first (rather than a plain label
 * search) keeps this safe even if Annual Overview ever gains another
 * section reusing that same label.
 */
function getLifestyleAnnualTarget() {
  const sheet = getFinanceSheet("Annual Overview");
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
 
  let anchorRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim().toLowerCase().includes(LIFESTYLE_ANNUAL_ESTIMATE_ANCHOR)) {
      anchorRow = i + 1;
      break;
    }
  }
  if (!anchorRow) return null;
 
  for (let r = anchorRow + 1; r <= lastRow; r++) {
    if (String(sheet.getRange(r, 1).getValue()).trim().toLowerCase() === "total variable costs") {
      return Number(sheet.getRange(r, 2).getValue()) || 0;
    }
  }
  return null;
}
 
/**
 * Aggregates every "Fixed - <Name>" pool's Annual Budget and spend so far
 * (Annual Budget − Remaining Balance) — these feed into Projected Annual
 * Savings just as much as Lifestyle/Mandatory do, but weren't previously
 * part of the daily snapshot at all.
 */
function getFixedPoolSnapshot() {
  const files = DriveApp.getFilesByName(FINANCE_SPREADSHEET_NAME);
  const ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.getActiveSpreadsheet();
 
  const pools = [];
  let totalBudget = 0;
  let totalSpent = 0;
 
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (!name.startsWith(FIXED_POOL_PREFIX)) return;
 
    const poolName = name.slice(FIXED_POOL_PREFIX.length);
    const annualBudget = Number(sheet.getRange(2, 1).getValue()) || 0;
    const remaining = Number(sheet.getRange(2, 2).getValue()) || 0;
    const spent = annualBudget - remaining;
 
    pools.push({ name: poolName, annualBudget, spent, percentUsed: pct(spent, annualBudget) });
    totalBudget += annualBudget;
    totalSpent += spent;
  });
 
  return { pools, totalBudget, totalSpent, percentUsed: pct(totalSpent, totalBudget) };
}
 
/** Rounds to one decimal place; returns null instead of dividing by zero. */
function pct(numerator, denominator) {
  if (!denominator) return null;
  return Math.round((numerator / denominator) * 1000) / 10;
}
 
/** How far through the calendar year today is, as a percentage — the pacing benchmark every "on track" claim should be measured against. */
function getYearProgressPercent() {
  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  const diffDays = Math.floor((now - start) / (1000 * 60 * 60 * 24)) + 1;
  const isLeap = (now.getFullYear() % 4 === 0 && now.getFullYear() % 100 !== 0) || now.getFullYear() % 400 === 0;
  return pct(diffDays, isLeap ? 366 : 365);
}
 
function buildFinanceSnapshot() {
  const now = new Date();
  const monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM");
 
  const lifestyleSheet = getFinanceSheet(LIFESTYLE_SHEET);
  const mandatorySheet = getFinanceSheet(MANDATORY_SHEET);
 
  const lifestyleSection = findMonthSection(lifestyleSheet, monthName);
  const mandatorySection = findMonthSection(mandatorySheet, monthName);
 
  const lifestyleTotal = lifestyleSection
    ? lifestyleSheet.getRange(lifestyleSection.totalRow, 2, 1, 3).getValues()[0]
    : null;
  const mandatoryTotal = mandatorySection
    ? mandatorySheet.getRange(mandatorySection.totalRow, 2, 1, 3).getValues()[0]
    : null;
 
  // Year-to-date rollups (maintained automatically by recomputeAnnualRollup
  // in financeCore.gs every time a transaction is applied).
  const lifestyleRollupCells = ANNUAL_ROLLUP_CELLS[LIFESTYLE_SHEET];
  const mandatoryRollupCells = ANNUAL_ROLLUP_CELLS[MANDATORY_SHEET];
  const lifestyleActualToDate = Number(lifestyleSheet.getRange(lifestyleRollupCells.actual).getValue()) || 0;
  const mandatoryActualToDate = Number(mandatorySheet.getRange(mandatoryRollupCells.actual).getValue()) || 0;
  const lifestyleAnnualTarget = getLifestyleAnnualTarget();
  const mandatoryAnnualTarget = getMandatoryAnnualTarget();
  const fixedPools = getFixedPoolSnapshot();
  const yearProgressPercent = getYearProgressPercent();
 
  return {
    monthName,
    lifestyle: lifestyleTotal ? { actual: lifestyleTotal[0], budget: lifestyleTotal[1], savings: lifestyleTotal[2] } : null,
    mandatory: mandatoryTotal ? { actual: mandatoryTotal[0], budget: mandatoryTotal[1], savings: mandatoryTotal[2] } : null,
    lifestyleActualToDate,
    mandatoryActualToDate,
    lifestyleAnnualTarget,
    mandatoryAnnualTarget,
    lifestylePercentUsed: pct(lifestyleActualToDate, lifestyleAnnualTarget),
    mandatoryPercentUsed: pct(mandatoryActualToDate, mandatoryAnnualTarget),
    fixedPools,
    yearProgressPercent,
    annualSavingsTarget: getAnnualSavingsTarget(),
    // Distinguishes "no month section exists yet" (nothing categorized) from
    // "a section exists and genuinely shows $0 actual" — collapsing these
    // into one produced misleading pace/progress commentary. Fixed pool
    // spend counts too, since that's real annual data even if the two
    // Monthly sheets haven't picked up anything yet.
    hasAnyDataThisYear: lifestyleActualToDate > 0 || mandatoryActualToDate > 0 || fixedPools.totalSpent > 0,
    unknownCount: getAllUnknownTransactions().length
  };
}
 
function generateFinanceEmailSummary(snapshot) {
  const lifestyleStatus = snapshot.lifestyle
    ? `actual: $${snapshot.lifestyle.actual}, budget: $${snapshot.lifestyle.budget}, difference (budget minus actual): $${snapshot.lifestyle.savings}`
    : "no transactions have been categorized to this sheet for this month yet";
  const mandatoryStatus = snapshot.mandatory
    ? `actual: $${snapshot.mandatory.actual}, budget: $${snapshot.mandatory.budget}, difference (budget minus actual): $${snapshot.mandatory.savings}`
    : "no transactions have been categorized to this sheet for this month yet";
 
  const fixedPoolLines = snapshot.fixedPools.pools.length > 0
    ? snapshot.fixedPools.pools.map(p =>
        `- ${p.name}: spent $${p.spent} of $${p.annualBudget} annual budget (${p.percentUsed ?? "unknown"}% used)`
      ).join("\n")
    : "(no Fixed pools found)";
 
  const prompt = `You are a personal finance assistant. Given this month's actual vs. budgeted spending, plus
year-to-date progress against annual targets — including Fixed pools (Travel, Medical, Car, etc.),
which are annual lump-sum budgets tracked separately from the monthly categories but still count
toward Projected Annual Savings — write a short, honest, encouraging-but-not-fluffy email
(150-250 words).
 
CRITICAL — read this before writing anything:
 
* "No transactions have been categorized to this sheet for this month yet" means exactly that:
  there is no spending data for this month, not that $0 was spent. Never restate this as "$0
  spent" or as evidence of good budgeting.
 
* hasAnyDataThisYear is ${snapshot.hasAnyDataThisYear}. If it is false, there is NO real
  year-to-date spending data at all yet (across Lifestyle, Mandatory, or Fixed pools). In that
  case do NOT mention percentages of the annual target, pace toward the savings goal, or claim
  anything is "on track" — there is nothing to base that on. Instead say plainly that tracking
  hasn't started yet this year and that once transactions are categorized, this email will start
  reflecting real progress.
 
* yearProgressPercent is ${snapshot.yearProgressPercent}% — this is what "on pace" actually means:
  a category at a higher %-used than this is running hot relative to the calendar, one at a lower
  %-used is running cold. Use this as the benchmark for any pacing claim instead of a vague
  impression — every percentUsed figure below is already computed, don't recompute or estimate
  your own.
 
* ${snapshot.unknownCount} transaction(s) are currently sitting uncategorized, waiting on a
  reply to be filed into a budget category (see the list below this email). Mention this only
  if the count is greater than 0, and frame it as the next action to take, not a problem.
 
If (and only if) hasAnyDataThisYear is true, cover:
 
* How this month's Lifestyle and Mandatory spending compared to budget (or that one/both aren't
  categorized yet, per the CRITICAL note above)
* Fixed pools: total spent vs. total annual budget and %-used, benchmarked against
  yearProgressPercent — and call out by name any individual pool running notably hot (%-used well
  above yearProgressPercent) or notably cold, if one stands out
* One specific thing that went well
* One specific, actionable thing to improve next month — this can point at a Monthly category or
  a specific Fixed pool, whichever the numbers actually support
* Tie it back to the annual savings target explicitly: given the combined pace across Lifestyle,
  Mandatory, and Fixed pools, is that pace broadly consistent with reaching
  Projected Annual Savings, or is something specific putting it at risk
 
This month (${snapshot.monthName}):
Lifestyle — ${lifestyleStatus}
Mandatory — ${mandatoryStatus}
 
Year-to-date (yearProgressPercent: ${snapshot.yearProgressPercent}% of the year elapsed):
Lifestyle spent so far: $${snapshot.lifestyleActualToDate}, annual Lifestyle target: $${snapshot.lifestyleAnnualTarget ?? "unknown"} (${snapshot.lifestylePercentUsed ?? "unknown"}% used)
Mandatory spent so far: $${snapshot.mandatoryActualToDate}, annual Mandatory target (excluding fixed-pool amounts): $${snapshot.mandatoryAnnualTarget ?? "unknown"} (${snapshot.mandatoryPercentUsed ?? "unknown"}% used)
 
Fixed pools — total spent: $${snapshot.fixedPools.totalSpent} of $${snapshot.fixedPools.totalBudget} annual budget (${snapshot.fixedPools.percentUsed ?? "unknown"}% used):
${fixedPoolLines}
 
Annual savings target: $${snapshot.annualSavingsTarget ?? "unknown"}
 
OUTPUT FORMAT:
 
Return ONLY a valid JSON object with exactly one property:
 
{
"body": "<p>HTML email body here...</p>"
}
 
BODY REQUIREMENTS:
 
* The "body" property must contain valid HTML suitable for inserting directly into an
  existing email template.
* Return only the content that belongs inside the email's existing content container.
* Do NOT return <html>, <head>, <body>, <!DOCTYPE>, <style>, <script>, or <meta> tags.
* Do NOT include CSS.
* Do NOT use Markdown.
* Use simple email-compatible HTML such as <p>, <strong>, <ul>, <li>, and <br>.
* Do not create the surrounding email template; the application will provide that.
* Keep the visible body approximately 150-250 words.
* Do not include a greeting unless it naturally fits the email.
* Do not include a sign-off unless it naturally fits the email.
* Do not include "Subject:" anywhere in the body.
 
IMPORTANT JSON REQUIREMENTS:
 
* Return valid JSON only.
* Do not wrap the JSON in Markdown code fences.
* Escape quotation marks inside the HTML when necessary so the result remains valid JSON.
* Do not include any properties other than "body".
`;
  return llm(prompt, { reasoningEffort: "low" });
}

/**
 * Deterministic (no LLM) HTML block listing every transaction currently
 * sitting in the Unknown bucket, for the email's rendered HTML body.
 */
function buildUnknownSectionText() {
  const unknowns = getAllUnknownTransactions();
  if (unknowns.length === 0) return "";

  const labels = listValidCategoryTargets()
    .map(t => t.targetSheet.startsWith("FixedPool:") ? t.targetSheet.split(":")[1] : (t.targetField || t.targetSheet))
    .filter(Boolean);

  const lines = unknowns.map(u => {
    const plaidCategory = u.rawCategory ? ` (Plaid: ${u.rawCategory})` : "";
    return `<li>[${u.transaction_id}] ${u.date} — ${u.name} — $${u.amount}${plaidCategory}</li>`;
  });

  const uncategorizedList = `<ul>${lines.join("")}</ul>`;
  const labelsList = `<ul>${labels.map(l => `<li>${l}</li>`).join("")}</ul>`;

  return `
<p><strong>UNCATEGORIZED TRANSACTIONS (${unknowns.length})</strong></p>
${uncategorizedList}
<p><strong>To categorize these, reply to this email with one line per transaction:</strong></p>
<p>&lt;transaction ID&gt;: &lt;category&gt;</p>
<p><strong>Valid categories:</strong></p>
${labelsList}
`;
}

/** Crude HTML-to-plain-text fallback for non-HTML email clients — display only, never re-parsed by anything (replies are parsed from HTML — see parseCategoryReply). */
function stripHtmlTags(html) {
  return html
    .replace(/<li>/gi, "\n- ")
    .replace(/<\/li>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

/**
 * Parses a reply's HTML into {transactionId, category} pairs. Reads the
 * raw HTML (not a separately-maintained plain-text copy) so there's only
 * one structure for the Unknown-transactions list to stay in sync — the
 * same buildUnknownSectionText() HTML that gets sent is what gets quoted
 * back in a reply and parsed here. Deliberately permissive on the category
 * text — validation against the exact-match whitelist happens downstream
 * in checkForCategoryReplies(), so a near-miss category gets reported back
 * to you instead of silently mis-filed by the LLM guessing the "closest" one.
 */
function parseCategoryReply(replyPlaintext, validLabels) {
  const prompt = `
You are parsing the plaintext body of an email reply that assigns finance transaction IDs to
categories. Only parse the person's own newly written text.

Transaction IDs in the original (quoted) email are written in square brackets, like [abc123].

Valid categories (for reference only — do not correct or guess-match, pass through what the
person actually wrote):
${validLabels.join(", ")}

Email reply PlainText:
"""
${replyPlaintext}
"""

Return ONLY a JSON array (no markdown, no preamble) of objects: {"transactionId": "...", "category": "..."}.
Only include entries that clearly reference a transaction ID from the original email.
`;
  const result = llm(prompt, { reasoningEffort: "low" });
  if (!result.success) return { success: false, error: result.error };

  try {
    const clean = result.data.replace(/```json|```/g, "").trim();
    const data = JSON.parse(clean);
    if (!Array.isArray(data)) return { success: false, error: "LLM did not return an array" };
    return { success: true, data };
  } catch (e) {
    return { success: false, error: "Failed to parse LLM JSON: " + e.toString() };
  }
}

/**
 * Sends the email as HTML (with an auto-derived plain-text fallback for
 * non-HTML clients — display only, not what replies get parsed from) and
 * returns the sent thread's ID. GmailApp.sendEmail doesn't hand back the
 * sent message, so look it up immediately after — safe since Gmail sends
 * are synchronous and this runs at most a few times/day, so there's no
 * ambiguity about which thread it is.
 */
function sendFinanceEmail(htmlBody, monthName) {
  const subject = `Finance Summary — ${monthName}`;
  GmailApp.sendEmail(MY_EMAIL, subject, stripHtmlTags(htmlBody), { htmlBody: htmlBody });

  const threads = GmailApp.search(`in:sent subject:"${subject}" to:${MY_EMAIL}`, 0, 1);
  return threads.length > 0 ? threads[0].getId() : null;
}

const FALLBACK_SUMMARY_HTML = "<p>This month's narrative summary couldn't be generated right now — the numbers and any uncategorized transactions are below.</p>";

/**
 * Calls generateFinanceEmailSummary and extracts the HTML body, retrying
 * once on failure (most commonly a truncated LLM response — finish_reason
 * "length" — which leaves the JSON unclosed and unparsable). Returns null
 * if both attempts fail, rather than throwing, so the caller can still
 * send the email with a fallback instead of dropping it entirely.
 */
function getSummaryHtmlWithRetry(snapshot, attempts) {
  attempts = attempts || 2;
  for (let i = 1; i <= attempts; i++) {
    const summary = generateFinanceEmailSummary(snapshot);
    if (!summary.success) {
      console.warn(`Summary attempt ${i}/${attempts} failed:`, summary.error);
      continue;
    }
    try {
      const clean = summary.data.replace(/```json|```/g, "").trim();
      return JSON.parse(clean).body;
    } catch (e) {
      console.warn(`Summary attempt ${i}/${attempts}: failed to parse LLM JSON (${e.message}) — likely a truncated response.`);
    }
  }
  console.error("All attempts to generate the finance summary failed — sending the email without the narrative.");
  return null;
}

/**
 * Builds and sends the finance summary email, and re-points (or clears)
 * FINANCE_REPLY_THREAD_ID based on whatever's left in the Unknown bucket.
 * Called by the daily cron (gated on hasChanges) and by
 * checkForCategoryReplies() (unconditionally, right after resolving
 * replies) so the sheets/email reflect new categorizations immediately
 * instead of waiting for the next day's run.
 *
 * The narrative summary is best-effort: if the LLM call fails or its
 * output is truncated, the email still sends with a fallback line rather
 * than being dropped — the Uncategorized Transactions list below it is
 * the part that actually needs to reach you every time.
 */
function sendFinanceSummaryEmail() {
  const snapshot = buildFinanceSnapshot();
  const summaryHtml = getSummaryHtmlWithRetry(snapshot) || FALLBACK_SUMMARY_HTML;

  const unknownSection = buildUnknownSectionText();
  const htmlBody = summaryHtml + unknownSection;
  const threadId = sendFinanceEmail(htmlBody, snapshot.monthName);

  if (unknownSection) {
    properties.setProperty('FINANCE_REPLY_THREAD_ID', threadId || "");
  } else {
    properties.deleteProperty('FINANCE_REPLY_THREAD_ID');
  }
}

/**
 * Entry point for the daily trigger: sync, and only email if the sync
 * actually changed something in the sheets.
 */
function dailyFinanceCheck() {
  ensureMonthSectionsUpToDate();
  const syncResult = runDailyFinanceSync();

  if (!syncResult.hasChanges) {
    console.log("No finance changes today — skipping email.");
    return;
  }

  sendFinanceSummaryEmail();

  if (syncResult.uncategorizedCount > 0) {
    console.warn(`${syncResult.uncategorizedCount} transaction(s) were uncategorized this run — check the Category Rules sheet.`);
  }
}

