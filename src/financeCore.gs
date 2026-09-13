// =========================================================================
// CONFIGURATION
// =========================================================================
const properties = PropertiesService.getScriptProperties();
const FINANCE_SPREADSHEET_NAME = properties.getProperty('FINANCE_SPREADSHEET_NAME');
const LIFESTYLE_SHEET = "Monthly LifeStyle Cost";
const MANDATORY_SHEET = "Monthly Mandatory Cost";
const LEDGER_SHEET = "Transaction Ledger";
const RULES_SHEET = "Category Rules";
const UNKNOWN_SHEET = "Uncategorized Transactions";
const UNKNOWN_TARGET = "Unknown";
const FIXED_POOL_PREFIX = "Fixed - ";
const FIXED_POOL_LEDGER_START_ROW = 5;
let financeSpreadsheetCache = null;

// Only transactions on/after this date get applied to the sheets — added,
// modified, AND removed transactions dated earlier are ignored entirely.
// Recomputed each execution (not a fixed literal), so it always means
// "January 1st of whatever year the script is running in."
const SYNC_START_DATE = `${new Date().getFullYear()}-01-01`;

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"];
const MONTH_NAMES_LOWER = MONTH_NAMES.map(m => m.toLowerCase());

/**
 * Opens (and caches, for the life of this execution) the Finances
 * spreadsheet by name. getFinanceSheet() was previously doing a fresh
 * DriveApp.getFilesByName() + SpreadsheetApp.open() on every single call —
 * and it's called many times per transaction (Ledger, Category Rules,
 * Monthly sheets, Unknown bucket...). That repeated Drive search was the
 * actual reason a single ~50-transaction sync page could take long enough
 * to exceed Apps Script's execution time limit on its own. Caching the
 * handle once per execution cuts that down to one Drive search per run.
 */
function getFinanceSpreadsheet() {
  if (financeSpreadsheetCache) return financeSpreadsheetCache;
  const files = DriveApp.getFilesByName(FINANCE_SPREADSHEET_NAME);
  financeSpreadsheetCache = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.getActiveSpreadsheet();
  return financeSpreadsheetCache;
}

/**
 * Opens a sheet by name within the Finances spreadsheet (a different
 * spreadsheet than the webscraper DB one — resolved separately by name,
 * same pattern as getTargetSheet in db.gs).
 */
function getFinanceSheet(sheetName) {
  const ss = getFinanceSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in "${FINANCE_SPREADSHEET_NAME}". Run setupFinanceSheets() first if this is a new sheet.`);
  }
  return sheet;
}

// =========================================================================
// CATEGORIZATION — driven entirely by the "Category Rules" sheet, since
// only the account owner knows which merchant/Plaid category maps to which
// custom bucket. Columns: Merchant Name Contains | Plaid Category (primary)
// Contains | Plaid Category (detailed) Contains | Target Sheet | Target Field.
// Target Sheet is one of: "Monthly LifeStyle Cost", "Monthly Mandatory Cost",
// "FixedPool:<PoolName>" (e.g. "FixedPool:Travel"), or "Ignore" — a
// transaction matching an Ignore rule is skipped entirely: not written to
// any budget sheet, not added to the Ledger, and not sent to the Unknown
// bucket. Built for cases like credit card payments, where the spend was
// already tracked when the original purchase posted and counting the
// payment too would double it.
//
// A rule matches when EVERY one of its non-blank match columns matches
// (AND, not OR) — leaving a column blank means "don't care" about that
// field. This lets the same merchant map to different categories depending
// on the Plaid-assigned category, and lets a specific detailed category
// (e.g. "loan_payments_credit_card_payment") be targeted without also
// catching other loan payments under the same primary category
// ("loan_payments" also covers mortgage/student/personal loan payments).
// Rules are checked top-to-bottom; first match wins, so put more specific
// (more-columns-filled) rules above broader ones.
// =========================================================================
const IGNORE_TARGET_SHEET = "Ignore";

function loadCategoryRules() {
  const sheet = getFinanceSheet(RULES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return values
    .filter(row => (row[0] || row[1] || row[2]) && row[3]) // at least one match column + a target sheet
    .map(row => ({
      merchantMatch: String(row[0] || "").trim().toLowerCase(),
      pfcMatch: String(row[1] || "").trim().toLowerCase(),
      pfcDetailedMatch: String(row[2] || "").trim().toLowerCase(),
      targetSheet: String(row[3]).trim(),
      targetField: String(row[4] || "").trim()
    }));
}

function categorizeTransaction(txn, rules) {
  const merchantName = String(txn.merchant_name || txn.name || "").toLowerCase();
  const pfcPrimary = String(txn.personal_finance_category?.primary || "").toLowerCase();
  const pfcDetailed = String(txn.personal_finance_category?.detailed || "").toLowerCase();

  for (const rule of rules) {
    const merchantOk = !rule.merchantMatch || merchantName.includes(rule.merchantMatch);
    const pfcOk = !rule.pfcMatch || pfcPrimary.includes(rule.pfcMatch);
    const pfcDetailedOk = !rule.pfcDetailedMatch || pfcDetailed.includes(rule.pfcDetailedMatch);
    if (merchantOk && pfcOk && pfcDetailedOk) {
      return { targetSheet: rule.targetSheet, targetField: rule.targetField };
    }
  }
  return null; // uncategorized — stored in the Unknown bucket, not guessed at
}

/**
 * Accepts either a plain "YYYY-MM-DD" string (what Plaid's API returns) or
 * a JS Date object — Google Sheets silently converts a date-looking string
 * into a real Date on write, so anything read back out of a sheet (e.g.
 * the Uncategorized Transactions sheet's Date column, consumed by
 * resolveUnknownTransaction) may come back as the latter even though it
 * was written as a string. Returns monthName: null on anything unparsable
 * instead of letting a bad date silently produce undefined further downstream.
 */
function getMonthYearFromDate(dateValue) {
  const dateStr = dateValue instanceof Date
    ? Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(dateValue);

  const [year, month] = dateStr.split("-");
  const monthIndex = parseInt(month, 10) - 1;
  const monthName = MONTH_NAMES[monthIndex] || null;
  return { monthName, year: parseInt(year, 10) };
}

// =========================================================================
// TRANSACTION LEDGER — the source of truth for "what's already been
// applied and where," so a modified/removed transaction can be reversed
// and reapplied correctly instead of just accumulating duplicate deltas.
// Columns: Transaction ID | Date | Amount | Description | Target Sheet |
//          Target Field | Month | Year
// =========================================================================
function getLedgerRowByTxnId(txnId) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === txnId) return i + 2;
  }
  return null;
}

function getLedgerRowValues(row) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const values = sheet.getRange(row, 1, 1, 8).getValues()[0];
  return {
    amount: values[2],
    targetSheet: values[4],
    targetField: values[5],
    monthName: values[6]
  };
}

function writeLedgerEntry(existingRow, txn, categorization, monthName, year) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const values = [
    txn.transaction_id,
    txn.date,
    txn.amount,
    txn.name || txn.merchant_name || "",
    categorization.targetSheet,
    categorization.targetField,
    monthName,
    year
  ];
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function deleteLedgerRow(row) {
  getFinanceSheet(LEDGER_SHEET).deleteRow(row);
}

// =========================================================================
// MONTHLY SHEET UPDATES (Monthly LifeStyle Cost / Monthly Mandatory Cost)
// Finds each month's section by scanning column A (rather than hardcoded
// row numbers), since the sheets have irregular layouts (e.g. an extra
// "Travel" row appears after some months but not others).
// =========================================================================
function findRowByLabelInRange(sheet, label, startRow, endRow) {
  if (endRow < startRow) return null;
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, 1).getValues();
  const target = label.trim().toLowerCase();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === target) return startRow + i;
  }
  return null;
}

function findMonthSection(sheet, monthName) {
  if (!monthName) return null; // guards against an upstream date-parsing failure instead of crashing here
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return null; // completely empty sheet — nothing to scan
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  let monthRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim().toLowerCase() === monthName.toLowerCase()) {
      monthRow = i + 1;
      break;
    }
  }
  if (!monthRow) return null;

  let totalRow = null;
  for (let r = monthRow + 1; r <= lastRow; r++) {
    const label = String(sheet.getRange(r, 1).getValue()).trim().toLowerCase();
    if (label === "total variable costs") { totalRow = r; break; }
    if (label && MONTH_NAMES_LOWER.includes(label)) break; // ran into the next month first
  }
  if (!totalRow) return null;
  return { monthRow, totalRow };
}

/** Finds every month section present in the sheet (not just one), for computing year-to-date rollups. */
function findAllMonthSections(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return []; // completely empty sheet — nothing to scan
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const sections = [];
  for (let i = 0; i < colA.length; i++) {
    const label = String(colA[i][0]).trim();
    if (MONTH_NAMES_LOWER.includes(label.toLowerCase())) {
      const section = findMonthSection(sheet, label);
      if (section) sections.push(Object.assign({ monthName: label }, section));
    }
  }
  return sections;
}

// Where each sheet's year-to-date rollup summary cells live (row 2, next to
// January): Monthly LifeStyle Cost has "Annual LifeStyle Cost" in F2 and
// "Savings" in G2; Monthly Mandatory Cost has "Annual Cost Ex Med & Car &
// Renters" in H2 and "Savings" in I2. These are recomputed as the sum of
// every month's actual (col B) and Budget-minus-Actual (col D) on the
// Total Variable Costs row — replacing the manual entry that was prone to
// drifting out of sync (e.g. one cell reflecting 2 months, the other 3).
const ANNUAL_ROLLUP_CELLS = {
  [LIFESTYLE_SHEET]: { actual: "F2", savings: "G2" },
  [MANDATORY_SHEET]: { actual: "H2", savings: "I2" }
};

function recomputeAnnualRollup(sheet, sheetName) {
  const cells = ANNUAL_ROLLUP_CELLS[sheetName];
  if (!cells) return; // sheet has no year-to-date rollup cells to maintain

  const sections = findAllMonthSections(sheet);
  let totalActual = 0;
  let totalSavings = 0;
  sections.forEach(section => {
    const [amount, , savings] = sheet.getRange(section.totalRow, 2, 1, 3).getValues()[0];
    totalActual += Number(amount) || 0;
    totalSavings += Number(savings) || 0;
  });

  sheet.getRange(cells.actual).setValue(totalActual);
  sheet.getRange(cells.savings).setValue(totalSavings);
}

// =========================================================================
// ANNUAL OVERVIEW AS SOURCE OF TRUTH — the two annual-estimate blocks in
// Annual Overview (Mandatory and Lifestyle) are the canonical list of
// categories + budgets. Monthly sections are built and kept in sync from
// these, rather than by copying an arbitrary earlier month.
// =========================================================================
const ANNUAL_OVERVIEW_SHEET = "Annual Overview";
const MANDATORY_SECTION_HEADER_ANCHOR = "mandatory non-discretionary annual costs";
const MANDATORY_FIXED_AMOUNTS_END_LABEL = "Total Annual Mandatory Fixed Amounts";
const TRAVEL_BUDGET_LABEL = "Total Annual Travel Budget";

/**
 * Reads every category/budget row between an anchor header (matched by
 * substring) and an end label (matched exactly) in column A of Annual
 * Overview, with the budget read from column B. Both the anchor row and
 * the end-label row are excluded. Used as the canonical source for a
 * Monthly sheet's categories and per-month budgets, and for the block of
 * individually-tracked Fixed-pool categories.
 */
function readAnnualOverviewCategoryBlock(anchorSubstring, endLabel) {
  const sheet = getFinanceSheet(ANNUAL_OVERVIEW_SHEET);
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();

  let anchorRow = null;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim().toLowerCase().includes(anchorSubstring)) {
      anchorRow = i + 1;
      break;
    }
  }
  if (!anchorRow) return [];

  const endTarget = endLabel.trim().toLowerCase();
  const categories = [];

  for (let r = anchorRow + 1; r <= lastRow; r++) {
    const label = String(sheet.getRange(r, 1).getValue()).trim();
    if (!label) continue;
    if (label.toLowerCase() === endTarget) break;
    const budget = Number(sheet.getRange(r, 2).getValue()) || 0;
    categories.push({ label, annualBudget: budget });
  }
  return categories;
}

function getLifestyleCategoryBudgets() {
  return readAnnualOverviewCategoryBlock(LIFESTYLE_ANNUAL_ESTIMATE_ANCHOR, "total variable costs");
}

/**
 * The per-month Mandatory budget categories — stops before the fixed-pool
 * block (Renters Insurance, Medical Expenses, Car Maintenance, Car
 * Registration, Car Insurance, etc.), which is read separately by
 * getMandatoryFixedPoolBudgets() and tracked as annual pools instead of
 * per-month line items.
 */
function getMandatoryCategoryBudgets() {
  return readAnnualOverviewCategoryBlock(MANDATORY_SECTION_HEADER_ANCHOR, MANDATORY_ANNUAL_TARGET_LABEL);
}

/**
 * The individually-tracked fixed-pool categories: every row between
 * "Total Mandatory Costs Excluding Fixed Amounts" and
 * "Total Annual Mandatory Fixed Amounts" — each becomes its own
 * "Fixed - <label>" pool (e.g. "Fixed - Medical Expenses"), synced by
 * syncFixedPoolsFromAnnualOverview().
 */
function getMandatoryFixedPoolBudgets() {
  return readAnnualOverviewCategoryBlock(MANDATORY_ANNUAL_TARGET_LABEL.toLowerCase(), MANDATORY_FIXED_AMOUNTS_END_LABEL);
}

function getCategoryBudgetsForSheet(sheetName) {
  if (sheetName === LIFESTYLE_SHEET) return getLifestyleCategoryBudgets();
  if (sheetName === MANDATORY_SHEET) return getMandatoryCategoryBudgets();
  return [];
}

/**
 * Creates (or updates) one "Fixed - <poolName>" pool's Annual Budget cell.
 * Never touches the pool's ledger rows — Remaining Balance is re-derived
 * from what's already spent whenever the budget figure changes, so a mid-
 * year Annual Overview edit doesn't erase spend history.
 */
function syncFixedPoolBudget(poolName, annualBudget) {
  if (annualBudget === null || annualBudget === undefined) return;
  const ss = getFinanceSpreadsheet();
  let sheet = null;
  for(let s of ss.getSheets()) {
    if(s.getName() === FIXED_POOL_PREFIX + poolName) sheet = s; return;
  };

  if (!sheet) {
    sheet = ss.insertSheet(FIXED_POOL_PREFIX + poolName);
    sheet.getRange(1, 1, 1, 2).setValues([["Annual Budget", "Remaining Balance"]]);
    sheet.getRange(4, 1, 1, 4).setValues([["Date", "Description", "Amount", "Transaction ID"]]);
    sheet.getRange(2, 1, 1, 2).setValues([[annualBudget, annualBudget]]);
    return;
  }

  const currentBudget = Number(sheet.getRange(2, 1).getValue()) || 0;
  if (currentBudget === annualBudget) return; // already in sync

  const lastRow = sheet.getLastRow();
  let spent = 0;
  if (lastRow >= FIXED_POOL_LEDGER_START_ROW) {
    const amounts = sheet.getRange(FIXED_POOL_LEDGER_START_ROW, 3, lastRow - FIXED_POOL_LEDGER_START_ROW + 1, 1).getValues();
    amounts.forEach(([amt]) => { spent += Number(amt) || 0; });
  }
  sheet.getRange(2, 1, 1, 2).setValues([[annualBudget, annualBudget - spent]]);
}

/**
 * Syncs every Annual-Overview-driven Fixed pool: one pool per row in the
 * Mandatory Fixed Amounts block (Renters Insurance, Medical Expenses, Car
 * Maintenance, Car Registration, Car Insurance, ...), plus Travel from its
 * own "Total Annual Travel Budget" line. Route transactions to any of
 * these via a Category Rule with Target Sheet = "FixedPool:<label exactly
 * as it appears in Annual Overview>" (e.g. "FixedPool:Medical Expenses").
 */
function syncFixedPoolsFromAnnualOverview() {
  getMandatoryFixedPoolBudgets().forEach(({ label, annualBudget }) => {
    syncFixedPoolBudget(label, annualBudget);
  });

  const travelBudget = getLabeledValue(getFinanceSheet(ANNUAL_OVERVIEW_SHEET), TRAVEL_BUDGET_LABEL);
  if (travelBudget !== null) syncFixedPoolBudget("Travel", travelBudget);
}

/**
 * Creates a brand-new month section for `monthName` in `sheet`, sourcing
 * its categories and budgets from `categories` (Annual Overview, via
 * getCategoryBudgetsForSheet) rather than copying an existing month.
 * Appended at the bottom — if you're backfilling months out of
 * chronological order, sections may not end up in calendar order; re-sort
 * manually afterward if that matters to you.
 */
function createMonthSection(sheet, monthName, categories) {
  if (!categories || categories.length === 0) {
    return { success: false, error: `No categories found in Annual Overview to build "${monthName}" from — check the annual-estimate block for "${sheet.getName()}" is populated.` };
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1).setValue(monthName);
  sheet.getRange(startRow, 1).setFontWeight("bold");

  let row = startRow + 1;
  let totalBudget = 0;
  categories.forEach(({ label, annualBudget }) => {
    const monthlyBudget = Math.round((annualBudget / 12) * 100) / 100;
    totalBudget += monthlyBudget;
    sheet.getRange(row, 1, 1, 4).setValues([[label, 0, monthlyBudget, monthlyBudget]]);
    sheet.getRange(row, 1, 1, 4).setFontWeight("normal");
    row++;
  });

  sheet.getRange(row, 1, 1, 4).setValues([["Total Variable Costs", 0, totalBudget, totalBudget]]);
  sheet.getRange(row, 1, 1, 4).setFontWeight("bold");

  return { success: true, section: { monthRow: startRow, totalRow: row } };
}

/**
 * Keeps an existing month section's category rows in sync with Annual
 * Overview: adds any missing category (Actual 0, Budget = annual/12), and
 * updates Budget (recomputing Savings) on any row whose budget no longer
 * matches. Never touches Actual — real spending data is never overwritten.
 * Mutates `section.totalRow` in place as rows are inserted.
 */
function reconcileMonthSectionWithCategories(sheet, section, categories) {
  let changed = false;
  categories.forEach(({ label, annualBudget }) => {
    const monthlyBudget = Math.round((annualBudget / 12) * 100) / 100;
    const existingRow = findRowByLabelInRange(sheet, label, section.monthRow + 1, section.totalRow - 1);
    if (existingRow) {
      const currentBudget = Number(sheet.getRange(existingRow, 3).getValue()) || 0;
      if (currentBudget !== monthlyBudget) {
        sheet.getRange(existingRow, 3).setValue(monthlyBudget);
        changed = true;
      }
    } else {
      // insertRowBefore copies formatting from the adjacent row — since
      // we're inserting directly above the (bold) Total row, the new row
      // would otherwise inherit that bold formatting. Explicitly reset it.
      sheet.insertRowBefore(section.totalRow);
      sheet.getRange(section.totalRow, 1, 1, 4).setValues([[label, 0, monthlyBudget, monthlyBudget]]);
      sheet.getRange(section.totalRow, 1, 1, 4).setFontWeight("normal");
      section.totalRow += 1;
      changed = true;
    }
  });
  if (changed) recomputeMonthTotal(sheet, section);
  return changed;
}

/**
 * Looks up a month section, creating it (via createMonthSection) if it
 * doesn't exist yet. Callable directly if you want to pre-create a month
 * ahead of any transactions landing in it.
 */
function ensureMonthSection(sheetName, monthName) {
  const sheet = getFinanceSheet(sheetName);
  const existing = findMonthSection(sheet, monthName);
  if (existing) return { success: true, section: existing, created: false };
  const created = createMonthSection(sheet, monthName, getCategoryBudgetsForSheet(sheetName));
  if (!created.success) return created;
  return { success: true, section: created.section, created: true };
}

/**
 * Entry point for the daily run: makes sure every month from January
 * through the current month has a section in both Monthly sheets
 * (building any missing one straight from Annual Overview), and
 * reconciles every existing section's categories/budgets against Annual
 * Overview so a category or budget change there propagates to every
 * month automatically. Also syncs the Medical/Car fixed pool.
 */
function ensureMonthSectionsUpToDate() {
  syncFixedPoolsFromAnnualOverview();

  const monthsToDate = MONTH_NAMES.slice(0, new Date().getMonth() + 1); // getMonth() is 0-indexed (0 = January)

  [
    { sheetName: LIFESTYLE_SHEET, categories: getLifestyleCategoryBudgets() },
    { sheetName: MANDATORY_SHEET, categories: getMandatoryCategoryBudgets() }
  ].forEach(({ sheetName, categories }) => {
    if (categories.length === 0) return; // Annual Overview block not populated yet — nothing to build from

    const sheet = getFinanceSheet(sheetName);
    monthsToDate.forEach(monthName => {
      const section = findMonthSection(sheet, monthName);
      if (!section) {
        createMonthSection(sheet, monthName, categories);
      } else {
        reconcileMonthSectionWithCategories(sheet, section, categories);
      }
    });

    recomputeAnnualRollup(sheet, sheetName);
  });
}

/** Adds `delta` to a field's Amount (actual) cell for the given month, then recomputes that month's total row and the sheet's year-to-date rollup. */
function adjustFieldAmount(sheetName, monthName, fieldLabel, delta) {
  if (!monthName) return { success: false, error: `Could not determine a valid month — check the transaction's date value.` };

  const sheet = getFinanceSheet(sheetName);
  let section = findMonthSection(sheet, monthName);
  if (!section) {
    const created = createMonthSection(sheet, monthName, getCategoryBudgetsForSheet(sheetName));
    if (!created.success) return created;
    section = created.section;
  }

  const fieldRow = findRowByLabelInRange(sheet, fieldLabel, section.monthRow + 1, section.totalRow - 1);
  if (!fieldRow) return { success: false, error: `Field "${fieldLabel}" not found under ${monthName} in ${sheetName}` };

  const amountCell = sheet.getRange(fieldRow, 2);
  const currentAmount = Number(amountCell.getValue()) || 0;
  amountCell.setValue(currentAmount + delta);

  recomputeMonthTotal(sheet, section);
  // NOTE: recomputeAnnualRollup() intentionally NOT called here — it does a
  // full column-A scan of the sheet (findAllMonthSections) plus a re-read
  // of every existing month section. Calling that once per transaction was
  // the main reason a large sync batch (e.g. a Jan-1 backfill) could take
  // long enough to hit Apps Script's execution time limit. It's now called
  // once per sheet at the end of runDailyFinanceSync() and
  // ensureMonthSectionsUpToDate() instead — everywhere that touches
  // multiple transactions in one run.
  return { success: true };
}

/** Recomputes a month's "Total Variable Costs" row: sum of Amount (col B), and Savings = Budget − Amount (col D). */
function recomputeMonthTotal(sheet, section) {
  const rows = sheet.getRange(section.monthRow + 1, 2, section.totalRow - section.monthRow - 1, 1).getValues();
  let totalAmount = 0;
  rows.forEach(([amount]) => { totalAmount += Number(amount) || 0; });

  sheet.getRange(section.totalRow, 2).setValue(totalAmount);
  const budget = Number(sheet.getRange(section.totalRow, 3).getValue()) || 0;
  sheet.getRange(section.totalRow, 4).setValue(budget - totalAmount);
}

// =========================================================================
// FIXED-POOL LEDGERS (Fixed - Travel, Fixed - Medical, etc.)
// Layout: row 2 = [Annual Budget, Remaining Balance]; ledger rows start at
// FIXED_POOL_LEDGER_START_ROW: Date | Description | Amount | Transaction ID.
// Remaining Balance is always recomputed from the ledger, not hand-edited.
// =========================================================================
function adjustFixedPool(poolName, txn, delta) {
  const sheet = getFinanceSheet(FIXED_POOL_PREFIX + poolName);
  const annualBudget = Number(sheet.getRange(2, 1).getValue()) || 0;

  const lastRow = sheet.getLastRow();
  let existingRow = null;
  if (lastRow >= FIXED_POOL_LEDGER_START_ROW) {
    const ids = sheet.getRange(FIXED_POOL_LEDGER_START_ROW, 4, lastRow - FIXED_POOL_LEDGER_START_ROW + 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === txn.transaction_id) { existingRow = FIXED_POOL_LEDGER_START_ROW + i; break; }
    }
  }

  if (delta === null) {
    // Removal
    if (existingRow) sheet.deleteRow(existingRow);
  } else if (existingRow) {
    sheet.getRange(existingRow, 1, 1, 3).setValues([[txn.date, txn.name || txn.merchant_name || "", txn.amount]]);
  } else {
    sheet.appendRow([txn.date, txn.name || txn.merchant_name || "", txn.amount, txn.transaction_id]);
  }

  const newLastRow = sheet.getLastRow();
  let spent = 0;
  if (newLastRow >= FIXED_POOL_LEDGER_START_ROW) {
    const amounts = sheet.getRange(FIXED_POOL_LEDGER_START_ROW, 3, newLastRow - FIXED_POOL_LEDGER_START_ROW + 1, 1).getValues();
    amounts.forEach(([amt]) => { spent += Number(amt) || 0; });
  }
  sheet.getRange(2, 2).setValue(annualBudget - spent);
}

// =========================================================================
// UNKNOWN BUCKET (Uncategorized Transactions) — for any transaction no
// Category Rule matches. Kept in sync via the same add/modify/remove
// lifecycle as everything else (see applyTransaction below), so a still-
// uncategorized transaction that later gets its amount corrected, or gets
// removed by the bank, stays accurate here too — not just frozen at
// whatever it looked like the first time it was seen.
// Columns: Transaction ID | Date | Description | Amount | Plaid Category
//          (detailed, for display) | Merchant Name | Plaid Category (primary)
//          | Plaid Category (detailed, raw)
// The last three are carried specifically so resolveUnknownTransaction()
// can auto-generate a precisely-matching Category Rule once you assign a
// category — including the detailed category, needed to target something
// like "credit card payment" without also catching every other kind of
// loan payment under the same primary category.
// =========================================================================
function addOrUpdateUnknownTransaction(txn) {
  const sheet = getFinanceSheet(UNKNOWN_SHEET);
  const lastRow = sheet.getLastRow();
  let existingRow = null;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === txn.transaction_id) { existingRow = i + 2; break; }
    }
  }
  const merchantName = txn.merchant_name || txn.name || "";
  const pfcPrimary = txn.personal_finance_category?.primary || "";
  const pfcDetailedRaw = txn.personal_finance_category?.detailed || "";
  const pfcDetailed = pfcDetailedRaw || pfcPrimary;
  const values = [txn.transaction_id, txn.date, txn.name || merchantName, txn.amount, pfcDetailed, merchantName, pfcPrimary, pfcDetailedRaw];
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function removeUnknownTransaction(txnId) {
  const sheet = getFinanceSheet(UNKNOWN_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === txnId) { sheet.deleteRow(i + 2); return; }
  }
}

function getUnknownTransaction(txnId) {
  const sheet = getFinanceSheet(UNKNOWN_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (const row of rows) {
    if (row[0] === txnId) {
      return {
        transaction_id: row[0], date: row[1], name: row[2], amount: row[3],
        rawCategory: row[4], merchant_name: row[5], pfc_primary: row[6], pfc_detailed: row[7]
      };
    }
  }
  return null;
}

function getAllUnknownTransactions() {
  const sheet = getFinanceSheet(UNKNOWN_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 8).getValues().map(row => ({
    transaction_id: row[0], date: row[1], name: row[2], amount: row[3],
    rawCategory: row[4], merchant_name: row[5], pfc_primary: row[6], pfc_detailed: row[7]
  }));
}

/**
 * Enumerates every valid category a transaction could be assigned to,
 * built dynamically from what's actually in the sheets (field labels
 * under any month's section, plus every "Fixed - <Name>" sheet), plus the
 * fixed "Ignore" option — used both for the email's reply instructions and
 * for the reply parser's exact-match whitelist.
 */
function listValidCategoryTargets() {
  const targets = [];

  [LIFESTYLE_SHEET, MANDATORY_SHEET].forEach(sheetName => {
    const sheet = getFinanceSheet(sheetName);
    const sections = findAllMonthSections(sheet);
    if (sections.length > 0) {
      const section = sections[0];
      const labels = sheet.getRange(section.monthRow + 1, 1, section.totalRow - section.monthRow - 1, 1).getValues();
      labels.forEach(([label]) => {
        if (label) targets.push({ targetSheet: sheetName, targetField: String(label).trim() });
      });
    }
  });

  const ss = getFinanceSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(FIXED_POOL_PREFIX)) {
      targets.push({ targetSheet: `FixedPool:${name.slice(FIXED_POOL_PREFIX.length)}`, targetField: "" });
    }
  });

  targets.push({ targetSheet: IGNORE_TARGET_SHEET, targetField: "" });

  return targets;
}

/**
 * Maps a human-facing category label (Monthly sheet field name, Fixed pool
 * name, or "Ignore") to its {targetSheet, targetField}. First-wins on
 * label collision across sheets — fine as long as field names stay
 * distinct, but worth knowing if you ever reuse a label like "Travel" in
 * both a Monthly sheet and a Fixed pool.
 */
function buildCategoryLabelMap(validTargets) {
  const map = {};
  (validTargets || listValidCategoryTargets()).forEach(t => {
    const label = t.targetSheet.startsWith("FixedPool:") ? t.targetSheet.split(":")[1] : (t.targetField || t.targetSheet);
    if (!label) return;
    const key = label.trim().toLowerCase();
    if (!map[key]) map[key] = t;
  });
  return map;
}

/**
 * Auto-learns a Category Rule from a resolved Unknown transaction, so the
 * same merchant + Plaid category combination auto-categorizes on future
 * syncs instead of landing in Unknown again. Matching on merchant + PFC
 * primary + PFC detailed (not merchant alone) means a repeat visit to the
 * same store under a different Plaid category — or even the same primary
 * category but a different detailed one, e.g. a loan payment vs. a credit
 * card payment, both under "loan_payments" — still falls through to
 * Unknown on its first occurrence and earns its own precise rule, rather
 * than silently inheriting whatever category was assigned last time.
 */
function addCategoryRuleIfMissing(merchantName, pfcPrimary, pfcDetailed, targetSheet, targetField) {
  if (!merchantName && !pfcPrimary && !pfcDetailed) return; // nothing distinctive to build a rule from

  const sheet = getFinanceSheet(RULES_SHEET);
  const lastRow = sheet.getLastRow();
  const m = String(merchantName || "").trim().toLowerCase();
  const p = String(pfcPrimary || "").trim().toLowerCase();
  const d = String(pfcDetailed || "").trim().toLowerCase();

  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const alreadyExists = existing.some(row =>
      String(row[0] || "").trim().toLowerCase() === m &&
      String(row[1] || "").trim().toLowerCase() === p &&
      String(row[2] || "").trim().toLowerCase() === d &&
      String(row[3] || "").trim() === targetSheet &&
      String(row[4] || "").trim() === targetField
    );
    if (alreadyExists) return;
  }

  sheet.appendRow([merchantName || "", pfcPrimary || "", pfcDetailed || "", targetSheet, targetField]);
}

/**
 * Assigns a transaction sitting in the Unknown bucket to a real category
 * (or to "Ignore") — called once a reply is parsed (see
 * checkForCategoryReplies in financeSchedule.gs). Applies it, records it,
 * removes it from Unknown, and auto-creates a matching Category Rule
 * (merchant + PFC primary + PFC detailed). Then sweeps the rest of the
 * Unknown bucket for any OTHER transaction with the exact same merchant +
 * PFC primary + PFC detailed and resolves/removes each of those too.
 * Returns the list of any additionally auto-resolved transactions
 * (excluding the one originally requested) so the caller can report them.
 */
function resolveUnknownTransaction(txnId, targetSheet, targetField) {
  const txn = getUnknownTransaction(txnId);
  if (!txn) {
    return { success: false, error: `No unknown transaction found for id ${txnId}` };
  }

  // Captured before the primary transaction is removed, and explicitly
  // excludes it by ID — otherwise it would match itself and get applied
  // twice.
  const otherMatches = getAllUnknownTransactions().filter(u =>
    u.transaction_id !== txn.transaction_id &&
    u.merchant_name === txn.merchant_name &&
    u.pfc_primary === txn.pfc_primary &&
    u.pfc_detailed === txn.pfc_detailed
  );

  const primaryResult = applyResolvedTransaction(txn, targetSheet, targetField);
  if (!primaryResult.success) return primaryResult;

  addCategoryRuleIfMissing(txn.merchant_name, txn.pfc_primary, txn.pfc_detailed, targetSheet, targetField);

  const alsoResolved = [];
  otherMatches.forEach(match => {
    const result = applyResolvedTransaction(match, targetSheet, targetField);
    if (result.success) {
      alsoResolved.push(match);
    } else {
      // A single bad-date secondary match shouldn't fail the whole reply —
      // the primary transaction the person actually replied about already
      // succeeded. Log it and leave that one in Unknown for next time.
      console.warn(`Auto-resolve skipped for ${match.transaction_id} (matched ${txn.transaction_id}'s new rule):`, result.error);
    }
  });

  // adjustFieldAmount() no longer recomputes the annual rollup per call
  // (see its comment) — do it once here instead, covering the primary
  // transaction and every swept-up match in one pass.
  if (!String(targetSheet).startsWith("FixedPool:") && targetSheet.toLowerCase() !== IGNORE_TARGET_SHEET.toLowerCase()) {
    recomputeAnnualRollup(getFinanceSheet(targetSheet), targetSheet);
  }

  return { success: true, alsoResolved };
}

/**
 * Applies a category (or "Ignore") to one Unknown-bucket transaction:
 * updates the target sheet (or Fixed pool), records it in the Ledger, and
 * removes it from Unknown — or, for "Ignore", just removes it from
 * Unknown without writing anywhere, since it's deliberately not tracked.
 * Shared by resolveUnknownTransaction() for both the originally-requested
 * transaction and every other transaction it auto-matches.
 */
function applyResolvedTransaction(txn, targetSheet, targetField) {
  if (String(targetSheet).trim().toLowerCase() === IGNORE_TARGET_SHEET.toLowerCase()) {
    removeUnknownTransaction(txn.transaction_id);
    return { success: true };
  }

  const { monthName, year } = getMonthYearFromDate(txn.date);
  if (!monthName) {
    return { success: false, error: `Could not determine month from stored date "${txn.date}" for transaction ${txn.transaction_id}.` };
  }

  if (String(targetSheet).startsWith("FixedPool:")) {
    adjustFixedPool(targetSheet.split(":")[1], txn, txn.amount);
  } else {
    const result = adjustFieldAmount(targetSheet, monthName, targetField, txn.amount);
    if (!result.success) return result;
  }

  const ledgerRow = getLedgerRowByTxnId(txn.transaction_id);
  writeLedgerEntry(ledgerRow, txn, { targetSheet, targetField }, monthName, year);
  removeUnknownTransaction(txn.transaction_id);
  return { success: true };
}

// =========================================================================
// APPLY A SINGLE TRANSACTION (added, modified, or removed)
// =========================================================================
function applyTransaction(txn, isRemoval) {
  const existingRow = getLedgerRowByTxnId(txn.transaction_id);
  const prevInfo = existingRow ? getLedgerRowValues(existingRow) : null;

  if (isRemoval) {
    if (prevInfo) {
      if (String(prevInfo.targetSheet).startsWith("FixedPool:")) {
        adjustFixedPool(prevInfo.targetSheet.split(":")[1], txn, null);
      } else {
        adjustFieldAmount(prevInfo.targetSheet, prevInfo.monthName, prevInfo.targetField, -Number(prevInfo.amount));
      }
      deleteLedgerRow(existingRow);
      return { changed: true };
    }
    // Not in the Ledger — it may be sitting in the Unknown bucket instead.
    if (getUnknownTransaction(txn.transaction_id)) {
      removeUnknownTransaction(txn.transaction_id);
      return { changed: true };
    }
    return { changed: false };
  }

  // Modified/added: reverse the previous effect first if this transaction
  // was already applied before (fixed pools handle their own overwrite by
  // transaction ID, so no separate reversal is needed there).
  if (prevInfo && !String(prevInfo.targetSheet).startsWith("FixedPool:")) {
    adjustFieldAmount(prevInfo.targetSheet, prevInfo.monthName, prevInfo.targetField, -Number(prevInfo.amount));
  }

  const rules = loadCategoryRules();
  const categorization = categorizeTransaction(txn, rules);
  if (!categorization) {
    if (existingRow) deleteLedgerRow(existingRow);
    addOrUpdateUnknownTransaction(txn); // dedupes by txn id itself
    return { changed: true, uncategorized: true };
  }

  // It matched a rule now — don't leave a stale copy in Unknown if it was
  // sitting there from a previous sync (e.g. a Category Rule was just added).
  removeUnknownTransaction(txn.transaction_id);

  if (String(categorization.targetSheet).trim().toLowerCase() === IGNORE_TARGET_SHEET.toLowerCase()) {
    // Deliberately not tracked (e.g. a credit card payment — the spend was
    // already counted when the original purchase posted). Any prior effect
    // was already reversed above; just drop the stale ledger row, if any,
    // and stop here — nothing gets written anywhere.
    if (existingRow) deleteLedgerRow(existingRow);
    return { changed: !!prevInfo }; // "changed" only if this reverses something that was previously tracked
  }

  const { monthName, year } = getMonthYearFromDate(txn.date);

  if (categorization.targetSheet.startsWith("FixedPool:")) {
    adjustFixedPool(categorization.targetSheet.split(":")[1], txn, txn.amount);
  } else {
    const result = adjustFieldAmount(categorization.targetSheet, monthName, categorization.targetField, txn.amount);
    if (!result.success) {
      console.warn("Could not apply transaction to sheet:", result.error);
      return { changed: false, uncategorized: false, error: result.error };
    }
  }

  writeLedgerEntry(existingRow, txn, categorization, monthName, year);
  return { changed: true };
}

// =========================================================================
// DAILY SYNC ORCHESTRATION
// =========================================================================
function runDailyFinanceSync() {
  const itemIds = JSON.parse(properties.getProperty('PLAID_ITEM_IDS') || "[]");
  let hasChanges = false;
  let uncategorizedCount = 0;

  // Apps Script kills execution around 6 minutes. Leave real margin so a
  // page in progress always finishes (and its cursor gets saved) rather
  // than getting killed mid-page — a large backlog (e.g. a Jan-1 backfill)
  // may take several daily runs to fully catch up, but each run is
  // guaranteed to make forward progress from wherever the last one's saved
  // cursor left off, instead of repeatedly dying at the same point.
  const startTime = Date.now();
  const TIME_BUDGET_MS = 5 * 60 * 1000; // 5 minutes

  for (const itemId of itemIds) {
    const accessToken = properties.getProperty(`PLAID_ACCESSS_TOKEN_${itemId}`);
    if (!accessToken) continue;

    // Captured once, before this item's pagination begins — the recovery
    // target if Plaid reports TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION
    // (the underlying data changed on Plaid's servers mid-pagination,
    // invalidating everything fetched so far in this session). Plaid's
    // guidance is to restart pagination from wherever this session started,
    // not from "" (a full resync from the beginning of history) and not
    // from wherever it broke.
    const sessionStartCursor = properties.getProperty(`PLAID_CURSOR_${itemId}`) || "";
    let currentCursor = sessionStartCursor;

    let hasMore = true;
    while (hasMore) {
      if (Date.now() - startTime > TIME_BUDGET_MS) {
        console.warn(`runDailyFinanceSync: time budget reached — stopping at item ${itemId}, will resume from the saved cursor next run.`);
        break;
      }

      const result = fetchNextSyncPage(accessToken, currentCursor);
      if (!result.success) {
        if (result.error === 'TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION') {
          console.warn(`${result.error} — restarting item ${itemId}'s sync from its session-starting cursor.`);
          currentCursor = sessionStartCursor;
          continue;
        }
        console.log(`Failed fetching for item ${itemId}: ${result.error}`);
        break;
      }
      const syncData = result.data;

      (syncData.added || []).forEach(txn => {
        if (txn.date < SYNC_START_DATE) return; // outside the tracked window — skip entirely
        const outcome = applyTransaction(txn, false);
        if (outcome.changed) hasChanges = true;
        if (outcome.uncategorized) uncategorizedCount++;
      });
      (syncData.modified || []).forEach(txn => {
        if (txn.date < SYNC_START_DATE) return;
        const outcome = applyTransaction(txn, false);
        if (outcome.changed) hasChanges = true;
        if (outcome.uncategorized) uncategorizedCount++;
      });
      (syncData.removed || []).forEach(removedTxn => {
        // Always attempt removal regardless of date — harmless no-op if it
        // was never applied (e.g. it was outside the window to begin with).
        const outcome = applyTransaction({ transaction_id: removedTxn.transaction_id }, true);
        if (outcome.changed) hasChanges = true;
      });

      // Advance BOTH the local cursor (so the next iteration of this same
      // execution actually moves forward) and the saved property (so the
      // next run resumes from here). Only updating the property while
      // leaving this local variable frozen was the actual bug behind
      // "same page reprocessed forever" and the resulting execution
      // timeout — the loop was refetching the identical page every
      // iteration since currentCursor never changed.
      currentCursor = syncData.next_cursor;
      properties.setProperty(`PLAID_CURSOR_${itemId}`, syncData.next_cursor);
      hasMore = syncData.has_more;
    }

    if (Date.now() - startTime > TIME_BUDGET_MS) break; // don't start a new item once time's up
  }

  // Recomputed once here instead of once per transaction inside
  // adjustFieldAmount() — see that function's comment for why.
  if (hasChanges) {
    [LIFESTYLE_SHEET, MANDATORY_SHEET].forEach(sheetName => {
      recomputeAnnualRollup(getFinanceSheet(sheetName), sheetName);
    });
  }

  return { hasChanges, uncategorizedCount };
}

// =========================================================================
// SWITCH TO A NEW SPREADSHEET — points the automation at a different
// Finances spreadsheet and clears each linked item's sync cursor so the
// next runDailyFinanceSync() re-pulls full available transaction history
// (filtered down to SYNC_START_DATE onward) instead of resuming where the
// old spreadsheet's cursor left off. Run this once manually, then either
// wait for the next daily trigger or call runDailyFinanceSync() /
// dailyFinanceCheck() directly to backfill immediately.
//
// Assumes the new spreadsheet already has Category Rules / Transaction
// Ledger / Uncategorized Transactions / Fixed pool sheets in place — if
// not, run setupFinanceSheets() first (after calling this, so it targets
// the new spreadsheet).
// =========================================================================
function startFreshFinanceSheet(newSpreadsheetName) {
  properties.setProperty('FINANCE_SPREADSHEET_NAME', newSpreadsheetName);

  const itemIds = JSON.parse(properties.getProperty('PLAID_ITEM_IDS') || "[]");
  itemIds.forEach(itemId => {
    properties.deleteProperty(`PLAID_CURSOR_${itemId}`);
  });

  properties.deleteProperty('FINANCE_REPLY_THREAD_ID');

  console.log(`Finance target switched to "${newSpreadsheetName}". Cursors cleared for ${itemIds.length} item(s) — the next sync will pull full available history, filtered to ${SYNC_START_DATE} onward.`);
}

// =========================================================================
// ONE-TIME SETUP — creates the new sheets this pipeline depends on, with
// headers, if they don't already exist. Run this once manually.
// =========================================================================
function setupFinanceSheets() {
  const ss = getFinanceSpreadsheet();

  if (!ss.getSheetByName(RULES_SHEET)) {
    const sheet = ss.insertSheet(RULES_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([[
      "Merchant Name Contains", "Plaid Category (primary) Contains", "Plaid Category (detailed) Contains", "Target Sheet", "Target Field"
    ]]);
    sheet.getRange(2, 1, 2, 5).setValues([
      ["", "groceries", "", "Monthly LifeStyle Cost", "Groceries"],
      ["", "loan_payments", "credit_card_payment", "Ignore", ""]
    ]);
  }

  if (!ss.getSheetByName(LEDGER_SHEET)) {
    const sheet = ss.insertSheet(LEDGER_SHEET);
    sheet.getRange(1, 1, 1, 8).setValues([[
      "Transaction ID", "Date", "Amount", "Description",
      "Target Sheet", "Target Field", "Month", "Year"
    ]]);
  }

  if (!ss.getSheetByName(UNKNOWN_SHEET)) {
    const sheet = ss.insertSheet(UNKNOWN_SHEET);
    sheet.getRange(1, 1, 1, 8).setValues([[
      "Transaction ID", "Date", "Description", "Amount", "Plaid Category", "Merchant Name", "Plaid Category (primary)", "Plaid Category (detailed)"
    ]]);
  }

  // Fixed pools (Travel + every category in the Mandatory Fixed Amounts
  // block) are no longer scaffolded here — syncFixedPoolsFromAnnualOverview()
  // creates each one automatically, sourcing its Annual Budget straight
  // from Annual Overview, the first time ensureMonthSectionsUpToDate() runs.

  console.log("Finance sheet setup complete. Fill in Category Rules and Annual Overview — Fixed pools are created automatically from it.");
}
