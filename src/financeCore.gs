// =========================================================================
// CONFIGURATION
// =========================================================================
const financeScriptProperties = PropertiesService.getScriptProperties();
const FINANCE_SPREADSHEET_NAME = financeScriptProperties.getProperty('FINANCE_SPREADSHEET_NAME');
const LIFESTYLE_SHEET = "Monthly LifeStyle Cost";
const MANDATORY_SHEET = "Monthly Mandatory Cost";
const LEDGER_SHEET = "Transaction Ledger";
const LEDGER_OVERRIDE_COLUMN = 9;
const LEDGER_DEFAULT_TARGET_SHEET_COLUMN = 10;
const LEDGER_DEFAULT_TARGET_FIELD_COLUMN = 11;
const UNKNOWN_CATEGORY_COLUMN = 9;
const CLEAR_TRANSACTION_OVERRIDE_LABEL = "Use original category";
const RULES_SHEET = "Category Rules";
const UNKNOWN_SHEET = "Uncategorized Transactions";
const UNKNOWN_TARGET = "Unknown";
const FIXED_POOL_PREFIX = "Fixed - ";
const FIXED_POOL_LEDGER_START_ROW = 5;
const PLAID_REPLACEMENT_REVIEW_PREFIX = "Plaid Review ";
const PLAID_REPLACEMENT_DATA_START_ROW = 7;
const PLAID_REPLACEMENT_ACTION_ROW = 5;
const PLAID_REPLACEMENT_ACTION_COLUMN = 2;
const PLAID_REPLACEMENT_FINALIZE_ACTION = "Finalize replacement";
const PLAID_REPLACEMENT_CANCEL_ACTION = "Cancel replacement";
let financeSpreadsheetCache = null;
let unknownCategoryValidationRuleCache = null;

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

function ensureSheetHasRows(sheet, requiredLastRow) {
  const maxRows = sheet.getMaxRows();
  if (requiredLastRow > maxRows) {
    sheet.insertRowsAfter(maxRows, Math.max(100, requiredLastRow - maxRows));
  }
}

function appendSheetRow(sheet, values) {
  const row = sheet.getLastRow() + 1;
  const previousMaxRows = sheet.getMaxRows();
  ensureSheetHasRows(sheet, row);
  if (sheet.getName() === LEDGER_SHEET && row > previousMaxRows) {
    sheet.getRange(previousMaxRows + 1, LEDGER_OVERRIDE_COLUMN, sheet.getMaxRows() - previousMaxRows, 1)
      .setDataValidation(getTransactionOverrideValidationRule());
  }
  if (sheet.getName() === UNKNOWN_SHEET) {
    setUnknownCategoryValidation(sheet, row, 1);
  }
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
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
      const fixedPoolName = getFixedPoolNameForTarget(rule.targetSheet);
      return {
        targetSheet: fixedPoolName ? `FixedPool:${fixedPoolName}` : rule.targetSheet,
        targetField: rule.targetField
      };
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
//          Target Field | Month | Year | Override Category |
//          Default Target Sheet | Default Target Field
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
  const values = sheet.getRange(row, 1, 1, 11).getValues()[0];
  return {
    amount: values[2],
    targetSheet: values[4],
    targetField: values[5],
    monthName: values[6],
    date: values[1],
    transactionId: values[0],
    description: values[3],
    overrideCategory: String(values[8] || ""),
    defaultTargetSheet: String(values[9] || ""),
    defaultTargetField: String(values[10] || "")
  };
}

function writeLedgerEntry(existingRow, txn, categorization, monthName, year, options) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const existingOverride = existingRow
    ? String(sheet.getRange(existingRow, LEDGER_OVERRIDE_COLUMN).getValue() || "")
    : "";
  const overrideCategory = options && Object.prototype.hasOwnProperty.call(options, "overrideCategory")
    ? options.overrideCategory
    : existingOverride;
  const defaultCategorization = options && options.defaultCategorization
    ? options.defaultCategorization
    : categorization;
  const values = [
    txn.transaction_id,
    txn.date,
    txn.amount,
    txn.name || txn.merchant_name || "",
    categorization.targetSheet,
    categorization.targetField,
    monthName,
    year,
    overrideCategory,
    defaultCategorization ? defaultCategorization.targetSheet : "",
    defaultCategorization ? defaultCategorization.targetField : ""
  ];
  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
  } else {
    appendSheetRow(sheet, values);
  }
}

function getTransactionOverrideTargetLabel(target) {
  const targetSheet = String(target.targetSheet || "").trim();
  if (targetSheet === IGNORE_TARGET_SHEET) return IGNORE_TARGET_SHEET;

  const fixedPoolName = getFixedPoolNameForTarget(targetSheet);
  if (fixedPoolName) return `Fixed — ${fixedPoolName}`;
  if (targetSheet === LIFESTYLE_SHEET) return `Lifestyle — ${target.targetField}`;
  if (targetSheet === MANDATORY_SHEET) return `Mandatory — ${target.targetField}`;
  throw new Error(`Unsupported transaction override target "${targetSheet}".`);
}

function getTransactionOverrideTargetMap() {
  const targets = listValidCategoryTargets();
  const targetMap = {};
  targets.forEach(target => {
    const label = getTransactionOverrideTargetLabel(target);
    targetMap[label.toLowerCase()] = target;
  });
  return targetMap;
}

function getTransactionOverrideValidationRule() {
  const targetMap = getTransactionOverrideTargetMap();
  const labels = [
    CLEAR_TRANSACTION_OVERRIDE_LABEL,
    ...Object.keys(targetMap).map(key => getTransactionOverrideTargetLabel(targetMap[key]))
  ];
  return SpreadsheetApp.newDataValidation()
    .requireValueInList([...new Set(labels)], true)
    .setAllowInvalid(false)
    .build();
}

function getUnknownCategoryValidationRule() {
  if (unknownCategoryValidationRuleCache) return unknownCategoryValidationRuleCache;
  const targetMap = getTransactionOverrideTargetMap();
  const labels = Object.keys(targetMap)
    .map(key => getTransactionOverrideTargetLabel(targetMap[key]));
  unknownCategoryValidationRuleCache = SpreadsheetApp.newDataValidation()
    .requireValueInList([...new Set(labels)], true)
    .setAllowInvalid(false)
    .build();
  return unknownCategoryValidationRuleCache;
}

function setUnknownCategoryValidation(sheet, startRow, numRows) {
  const ss = getFinanceSpreadsheet();
  if (!ss.getSheetByName(LIFESTYLE_SHEET) || !ss.getSheetByName(MANDATORY_SHEET)) return;
  sheet.getRange(startRow, UNKNOWN_CATEGORY_COLUMN, numRows, 1)
    .setDataValidation(getUnknownCategoryValidationRule());
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
const ACTING_BUDGET_LABEL = "Total Annual Acting Budget";
const SINKING_FUND_LABEL = "Sinking Fund";

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
  let sheet = ss.getSheetByName(FIXED_POOL_PREFIX + poolName);
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
  const actingBudget = getLabeledValue(getFinanceSheet(ANNUAL_OVERVIEW_SHEET), ACTING_BUDGET_LABEL);
  const sinkingFund = getLabeledValue(getFinanceSheet(ANNUAL_OVERVIEW_SHEET), SINKING_FUND_LABEL);
  
  if (travelBudget !== null) syncFixedPoolBudget("Travel", travelBudget);
  if (actingBudget !== null) syncFixedPoolBudget("Acting", actingBudget);
  if (sinkingFund !== null) syncFixedPoolBudget("Sinking", sinkingFund);
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
  const lastSectionRow = startRow + categories.length + 1;
  ensureSheetHasRows(sheet, lastSectionRow);
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
      ensureSheetHasRows(sheet, section.totalRow + 1);
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

function ensureMonthlyFinanceSheets(spreadsheet) {
  [LIFESTYLE_SHEET, MANDATORY_SHEET].forEach(sheetName => {
    if (!spreadsheet.getSheetByName(sheetName)) {
      spreadsheet.insertSheet(sheetName);
    }
  });
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
  ensureMonthlyFinanceSheets(getFinanceSpreadsheet());
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
    appendSheetRow(sheet, [txn.date, txn.name || txn.merchant_name || "", txn.amount, txn.transaction_id]);
  }

  const newLastRow = sheet.getLastRow();
  let spent = 0;
  if (newLastRow >= FIXED_POOL_LEDGER_START_ROW) {
    const amounts = sheet.getRange(FIXED_POOL_LEDGER_START_ROW, 3, newLastRow - FIXED_POOL_LEDGER_START_ROW + 1, 1).getValues();
    amounts.forEach(([amt]) => { spent += Number(amt) || 0; });
  }
  sheet.getRange(2, 2).setValue(annualBudget - spent);
}

function getFixedPoolNameForTarget(targetSheet) {
  const target = String(targetSheet || "").trim();
  const fixedPoolPrefix = "FixedPool:";
  if (target.startsWith(fixedPoolPrefix)) {
    const poolName = target.slice(fixedPoolPrefix.length).trim();
    return poolName || null;
  }
  if (target.startsWith(FIXED_POOL_PREFIX)) {
    const poolName = target.slice(FIXED_POOL_PREFIX.length).trim();
    return poolName || null;
  }

  return getFinanceSpreadsheet().getSheetByName(FIXED_POOL_PREFIX + target)
    ? target
    : null;
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
//          | Plaid Category (detailed, raw) | Assign to Category
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
    appendSheetRow(sheet, values);
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
 * fixed "Ignore" option — used to populate the sheet assignment dropdown.
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

  appendSheetRow(sheet, [merchantName || "", pfcPrimary || "", pfcDetailed || "", targetSheet, targetField]);
}

/**
 * Assigns a transaction sitting in the Unknown bucket to a real category
 * (or to "Ignore") — called when the sheet dropdown is selected. Applies
 * it, records it,
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
      // A single bad-date secondary match shouldn't fail the whole assignment —
      // the primary transaction the person actually replied about already
      // succeeded. Log it and leave that one in Unknown for next time.
      console.warn(`Auto-resolve skipped for ${match.transaction_id} (matched ${txn.transaction_id}'s new rule):`, result.error);
    }
  });

  // adjustFieldAmount() no longer recomputes the annual rollup per call
  // (see its comment) — do it once here instead, covering the primary
  // transaction and every swept-up match in one pass.
  if (!getFixedPoolNameForTarget(targetSheet) && targetSheet.toLowerCase() !== IGNORE_TARGET_SHEET.toLowerCase()) {
    recomputeAnnualRollup(getFinanceSheet(targetSheet), targetSheet);
  }

  return { success: true, alsoResolved };
}

function assignUnknownTransactionFromSheet(txnId, selectedLabel) {
  const target = getTransactionOverrideTargetMap()[String(selectedLabel || "").trim().toLowerCase()];
  if (!target) {
    throw new Error(`"${selectedLabel}" is not a valid category. Choose a value from the dropdown.`);
  }

  const outcome = resolveUnknownTransaction(txnId, target.targetSheet, target.targetField);
  if (!outcome.success) throw new Error(outcome.error);
  return outcome;
}

function recomputeAnnualRollupsForTargets(targets) {
  const monthlySheets = new Set();
  targets.forEach(target => {
    if (target.targetSheet === LIFESTYLE_SHEET || target.targetSheet === MANDATORY_SHEET) {
      monthlySheets.add(target.targetSheet);
    }
  });
  monthlySheets.forEach(sheetName => recomputeAnnualRollup(getFinanceSheet(sheetName), sheetName));
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

  const fixedPoolName = getFixedPoolNameForTarget(targetSheet);
  if (fixedPoolName) {
    adjustFixedPool(fixedPoolName, txn, txn.amount);
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
      removeTransactionTargetEffect(txn, prevInfo);
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

  const rules = loadCategoryRules();
  const ruleCategorization = categorizeTransaction(txn, rules);
  const defaultCategorization = ruleCategorization || (prevInfo && prevInfo.defaultTargetSheet
    ? { targetSheet: prevInfo.defaultTargetSheet, targetField: prevInfo.defaultTargetField }
    : null);
  const overrideCategory = prevInfo ? prevInfo.overrideCategory : "";
  const overrideTarget = overrideCategory
    ? getTransactionOverrideTargetMap()[overrideCategory.toLowerCase()]
    : null;
  if (overrideCategory && !overrideTarget) {
    return { changed: false, uncategorized: false, error: `Stored override "${overrideCategory}" is no longer a valid target for transaction ${txn.transaction_id}.` };
  }
  const categorization = overrideTarget || ruleCategorization;

  if (!categorization) {
    if (prevInfo) removeTransactionTargetEffect(txn, prevInfo);
    if (existingRow) deleteLedgerRow(existingRow);
    addOrUpdateUnknownTransaction(txn); // dedupes by txn id itself
    return { changed: true, uncategorized: true };
  }

  const { monthName, year } = getMonthYearFromDate(txn.date);
  if (!monthName) {
    return { changed: false, uncategorized: false, error: `Could not determine month from transaction date "${txn.date}" for transaction ${txn.transaction_id}.` };
  }

  if (String(categorization.targetSheet).trim().toLowerCase() !== IGNORE_TARGET_SHEET.toLowerCase()) {
    try {
      validateTransactionTarget(categorization, monthName);
    } catch (error) {
      const detail = error && error.message ? error.message : String(error);
      return { changed: false, uncategorized: false, error: detail };
    }
  }

  if (prevInfo) removeTransactionTargetEffect(txn, prevInfo);

  if (String(categorization.targetSheet).trim().toLowerCase() === IGNORE_TARGET_SHEET.toLowerCase()) {
    if (overrideTarget) {
      writeLedgerEntry(existingRow, txn, categorization, monthName, year, {
        overrideCategory,
        defaultCategorization
      });
      removeUnknownTransaction(txn.transaction_id);
      return { changed: true };
    }
    if (existingRow) deleteLedgerRow(existingRow);
    return { changed: !!prevInfo };
  }

  const result = addTransactionTargetEffect(txn, categorization, monthName);
  if (!result.success) {
    if (prevInfo) {
      try {
        addTransactionTargetEffect(txn, prevInfo, prevInfo.monthName);
      } catch (rollbackError) {
        const rollbackDetail = rollbackError && rollbackError.message
          ? rollbackError.message
          : String(rollbackError);
        return {
          changed: false,
          uncategorized: false,
          error: `${result.error} Rollback also failed: ${rollbackDetail}`
        };
      }
    }
    return { changed: false, uncategorized: false, error: result.error };
  }

  removeUnknownTransaction(txn.transaction_id);
  writeLedgerEntry(existingRow, txn, categorization, monthName, year, {
    overrideCategory,
    defaultCategorization
  });
  return { changed: true };
}

function removeTransactionTargetEffect(txn, previousTarget) {
  const previousFixedPoolName = getFixedPoolNameForTarget(previousTarget.targetSheet);
  if (previousFixedPoolName) {
    adjustFixedPool(previousFixedPoolName, txn, null);
  } else if (String(previousTarget.targetSheet).trim().toLowerCase() !== IGNORE_TARGET_SHEET.toLowerCase()) {
    const result = adjustFieldAmount(
      previousTarget.targetSheet,
      previousTarget.monthName,
      previousTarget.targetField,
      -Number(previousTarget.amount)
    );
    if (!result.success) throw new Error(result.error);
  }
}

function addTransactionTargetEffect(txn, target, monthName) {
  const fixedPoolName = getFixedPoolNameForTarget(target.targetSheet);
  if (fixedPoolName) {
    adjustFixedPool(fixedPoolName, txn, txn.amount);
    return { success: true };
  }
  if (String(target.targetSheet).trim().toLowerCase() === IGNORE_TARGET_SHEET.toLowerCase()) {
    return { success: true };
  }
  return adjustFieldAmount(target.targetSheet, monthName, target.targetField, Number(txn.amount));
}

function getTransactionTargetKey(target) {
  const fixedPoolName = getFixedPoolNameForTarget(target.targetSheet);
  if (fixedPoolName) return `fixed:${fixedPoolName.toLowerCase()}`;
  return `${String(target.targetSheet || "").trim().toLowerCase()}:${String(target.targetField || "").trim().toLowerCase()}`;
}

function validateTransactionTarget(target, monthName) {
  const targetSheet = String(target.targetSheet || "").trim();
  if (targetSheet.toLowerCase() === IGNORE_TARGET_SHEET.toLowerCase()) return;

  const fixedPoolName = getFixedPoolNameForTarget(targetSheet);
  if (fixedPoolName) {
    getFinanceSheet(FIXED_POOL_PREFIX + fixedPoolName);
    return;
  }

  if (targetSheet !== LIFESTYLE_SHEET && targetSheet !== MANDATORY_SHEET) {
    throw new Error(`Unsupported target sheet "${targetSheet}".`);
  }
  const section = ensureMonthSection(targetSheet, monthName);
  if (!section.success) throw new Error(section.error);
  const sheet = getFinanceSheet(targetSheet);
  const fieldRow = findRowByLabelInRange(
    sheet,
    target.targetField,
    section.section.monthRow + 1,
    section.section.totalRow - 1
  );
  if (!fieldRow) throw new Error(`Field "${target.targetField}" is not available in ${monthName} on ${targetSheet}.`);
}

function getTransactionOverrideChoiceForCurrentTarget(currentTarget, defaultTarget) {
  if (defaultTarget && getTransactionTargetKey(currentTarget) === getTransactionTargetKey(defaultTarget)) {
    return "";
  }
  return getTransactionOverrideTargetLabel(currentTarget);
}

function restoreTransactionLedgerOverrideSelection(row, previousValue) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const values = sheet.getRange(row, 5, 1, 7).getValues()[0];
  const inferredValue = getTransactionOverrideChoiceForCurrentTarget(
    { targetSheet: String(values[0] || ""), targetField: String(values[1] || "") },
    values[5]
      ? { targetSheet: String(values[5]), targetField: String(values[6] || "") }
      : null
  );
  sheet.getRange(row, LEDGER_OVERRIDE_COLUMN)
    .setValue(previousValue === undefined ? inferredValue : previousValue);
}

function processTransactionLedgerOverrideEdit(row, selectedLabel) {
  const sheet = getFinanceSheet(LEDGER_SHEET);
  const rowValues = sheet.getRange(row, 1, 1, 11).getValues()[0];
  const txn = {
    transaction_id: rowValues[0],
    date: rowValues[1],
    amount: Number(rowValues[2]),
    name: rowValues[3]
  };
  const previousTarget = {
    targetSheet: String(rowValues[4] || ""),
    targetField: String(rowValues[5] || ""),
    monthName: String(rowValues[6] || ""),
    amount: Number(rowValues[2])
  };
  const defaultTarget = rowValues[9]
    ? { targetSheet: String(rowValues[9]), targetField: String(rowValues[10] || "") }
    : previousTarget.targetSheet
      ? { targetSheet: previousTarget.targetSheet, targetField: previousTarget.targetField }
      : null;
  if (!defaultTarget) {
    throw new Error(`Transaction ${txn.transaction_id} has no stored target to use as its original category.`);
  }
  const label = String(selectedLabel || "").trim();
  const clearingOverride = !label || label.toLowerCase() === CLEAR_TRANSACTION_OVERRIDE_LABEL.toLowerCase();
  const destination = clearingOverride
    ? defaultTarget
    : getTransactionOverrideTargetMap()[label.toLowerCase()];
  if (!destination) {
    throw new Error(clearingOverride
      ? "This transaction has no stored original category to restore."
      : `"${label}" is not a valid override category. Choose a value from the dropdown.`);
  }
  const { monthName, year } = getMonthYearFromDate(txn.date);
  if (!monthName || !Number.isFinite(year)) {
    throw new Error(`Could not determine the month and year from transaction date "${txn.date}".`);
  }
  validateTransactionTarget(destination, monthName);
  validateTransactionTarget(previousTarget, previousTarget.monthName || monthName);

  const destinationCategory = clearingOverride ? "" : getTransactionOverrideTargetLabel(destination);
  if (getTransactionTargetKey(previousTarget) !== getTransactionTargetKey(destination)) {
    removeTransactionTargetEffect(txn, previousTarget);
    let result;
    try {
      result = addTransactionTargetEffect(txn, destination, monthName);
    } catch (error) {
      try {
        addTransactionTargetEffect(txn, previousTarget, previousTarget.monthName);
      } catch (rollbackError) {
        const rollbackDetail = rollbackError && rollbackError.message
          ? rollbackError.message
          : String(rollbackError);
        const detail = error && error.message ? error.message : String(error);
        throw new Error(`${detail} Rollback also failed: ${rollbackDetail}`);
      }
      throw error;
    }
    if (!result.success) {
      try {
        addTransactionTargetEffect(txn, previousTarget, previousTarget.monthName);
      } catch (rollbackError) {
        const rollbackDetail = rollbackError && rollbackError.message
          ? rollbackError.message
          : String(rollbackError);
        throw new Error(`${result.error} Rollback also failed: ${rollbackDetail}`);
      }
      throw new Error(result.error);
    }
  }

  sheet.getRange(row, 5, 1, 2).setValues([[destination.targetSheet, destination.targetField || ""]]);
  sheet.getRange(row, LEDGER_OVERRIDE_COLUMN).setValue(destinationCategory);
  sheet.getRange(row, LEDGER_DEFAULT_TARGET_SHEET_COLUMN, 1, 2).setValues([[
    defaultTarget.targetSheet,
    defaultTarget.targetField || ""
  ]]);

  recomputeAnnualRollupsForTargets([previousTarget, destination]);
  getFinanceSpreadsheet().toast(
    clearingOverride
      ? `Transaction restored to ${getTransactionOverrideTargetLabel(destination)}.`
      : `Transaction moved to ${getTransactionOverrideTargetLabel(destination)}.`,
    "Transaction category updated",
    5
  );
}

// =========================================================================
// DAILY SYNC ORCHESTRATION
// =========================================================================
function runDailyFinanceSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.warn("runDailyFinanceSync: another sync or replacement operation is active; this run will be retried.");
    return { hasChanges: false, uncategorizedCount: 0, needsContinuation: true, startedFromBeginning: false, errors: [] };
  }

  try {
    return runDailyFinanceSyncLocked();
  } finally {
    lock.releaseLock();
  }
}

function runDailyFinanceSyncLocked() {
  const itemIds = JSON.parse(financeScriptProperties.getProperty('PLAID_ITEM_IDS') || "[]");
  let hasChanges = false;
  let uncategorizedCount = 0;
  let needsContinuation = false;
  let startedFromBeginning = false;
  const errors = [];

  // Apps Script kills execution around 6 minutes. Leave real margin so a
  // page in progress always finishes (and its cursor gets saved) rather
  // than getting killed mid-page — a large backlog (e.g. a Jan-1 backfill)
  // may take several daily runs to fully catch up, but each run is
  // guaranteed to make forward progress from wherever the last one's saved
  // cursor left off, instead of repeatedly dying at the same point.
  const startTime = Date.now();
  const TIME_BUDGET_MS = 5 * 60 * 1000; // 5 minutes

  for (const itemId of itemIds) {
    const accessToken = financeScriptProperties.getProperty(`PLAID_ACCESS_TOKEN_${itemId}`);
    if (!accessToken) {
      const message = `Missing access token for active Plaid Item ${itemId}.`;
      console.error(`runDailyFinanceSync: ${message}`);
      errors.push(message);
      needsContinuation = true;
      continue;
    }

    // Captured once, before this item's pagination begins — the recovery
    // target if Plaid reports TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION
    // (the underlying data changed on Plaid's servers mid-pagination,
    // invalidating everything fetched so far in this session). Plaid's
    // guidance is to restart pagination from wherever this session started,
    // not from "" (a full resync from the beginning of history) and not
    // from wherever it broke.
    const savedCursor = financeScriptProperties.getProperty(`PLAID_CURSOR_${itemId}`);
    if (!savedCursor) startedFromBeginning = true;
    const sessionStartCursor = savedCursor || "";
    let currentCursor = sessionStartCursor;

    let hasMore = true;
    while (hasMore) {
      if (Date.now() - startTime > TIME_BUDGET_MS) {
        console.warn(`runDailyFinanceSync: time budget reached — stopping at item ${itemId}, will resume from the saved cursor next run.`);
        needsContinuation = true;
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
        errors.push(`Plaid sync failed for Item ${itemId}: ${result.error}`);
        needsContinuation = true;
        break;
      }
      const syncData = result.data;

      try {
        for (const txn of syncData.added || []) {
          if (txn.date < SYNC_START_DATE) continue;
          const outcome = applyTransaction(txn, false);
          if (outcome.error) throw new Error(outcome.error);
          if (outcome.changed) hasChanges = true;
          if (outcome.uncategorized) uncategorizedCount++;
        }
        for (const txn of syncData.modified || []) {
          if (txn.date < SYNC_START_DATE) continue;
          const outcome = applyTransaction(txn, false);
          if (outcome.error) throw new Error(outcome.error);
          if (outcome.changed) hasChanges = true;
          if (outcome.uncategorized) uncategorizedCount++;
        }
        for (const removedTxn of syncData.removed || []) {
          const outcome = applyTransaction({ transaction_id: removedTxn.transaction_id }, true);
          if (outcome.error) throw new Error(outcome.error);
          if (outcome.changed) hasChanges = true;
        }
      } catch (error) {
        const detail = error && error.message ? error.message : String(error);
        const message = `Could not apply transactions for Plaid Item ${itemId}: ${detail}`;
        console.error(message);
        errors.push(message);
        needsContinuation = true;
        break;
      }

      // Advance BOTH the local cursor (so the next iteration of this same
      // execution actually moves forward) and the saved property (so the
      // next run resumes from here). Only updating the property while
      // leaving this local variable frozen was the actual bug behind
      // "same page reprocessed forever" and the resulting execution
      // timeout — the loop was refetching the identical page every
      // iteration since currentCursor never changed.
      currentCursor = syncData.next_cursor;
      financeScriptProperties.setProperty(`PLAID_CURSOR_${itemId}`, syncData.next_cursor);
      hasMore = syncData.has_more;
    }

    if (hasMore) needsContinuation = true;
    if (Date.now() - startTime > TIME_BUDGET_MS) {
      needsContinuation = true;
      break;
    }
  }

  // Recomputed once here instead of once per transaction inside
  // adjustFieldAmount() — see that function's comment for why.
  if (hasChanges) {
    [LIFESTYLE_SHEET, MANDATORY_SHEET].forEach(sheetName => {
      recomputeAnnualRollup(getFinanceSheet(sheetName), sheetName);
    });
  }

  return { hasChanges, uncategorizedCount, needsContinuation, startedFromBeginning, errors };
}

function getPlaidReplacementReviewSheetName(itemId) {
  const safeItemId = String(itemId).replace(/[\[\]*\/\\?:]/g, "_").slice(0, 80);
  return `${PLAID_REPLACEMENT_REVIEW_PREFIX}${safeItemId}`;
}

function getPlaidReplacementReviewSheet(itemId) {
  const ss = getFinanceSpreadsheet();
  const sheetName = getPlaidReplacementReviewSheetName(itemId);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet && String(sheet.getRange(2, 2).getValue()) !== String(itemId)) {
    throw new Error(`Review sheet name collision for Plaid Item ${itemId}: "${sheetName}" belongs to another Item.`);
  }

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 5, 2).setValues([
      ["Old Item to replace (select bank / Item ID)", ""],
      ["New Item ID", itemId],
      ["Staging status", "Ready to stage transactions"],
      ["Instructions", ""],
      ["Replacement action", ""]
    ]);
    sheet.getRange(6, 1, 1, 11).setValues([[
      "Decision", "Status", "Transaction ID", "Date", "Amount", "Description",
      "Merchant Name", "PFC Primary", "PFC Detailed", "Account ID",
      "Possible Existing Transaction IDs"
    ]]);
    sheet.setFrozenRows(6);
  }
  sheet.getRange(PLAID_REPLACEMENT_ACTION_ROW, 1).setValue("Replacement action");
  sheet.getRange(PLAID_REPLACEMENT_ACTION_ROW, PLAID_REPLACEMENT_ACTION_COLUMN)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([
        PLAID_REPLACEMENT_FINALIZE_ACTION,
        PLAID_REPLACEMENT_CANCEL_ACTION
      ], true)
      .setAllowInvalid(false)
      .build())
    .setNote("Finalize imports the staged decisions and replaces the selected active Item. Cancel deletes this review sheet and pending local Item credentials; cancellation is blocked after any transactions have been imported or replaced.");
  ensureTransactionLedgerEditTrigger(ss);
  const oldItemIds = JSON.parse(financeScriptProperties.getProperty("PLAID_ITEM_IDS") || "[]");
  const oldItemOptions = oldItemIds.map(oldItemId => {
    const bankName = getPlaidItemInstitutionName(oldItemId);
    return `${bankName} — ${oldItemId}`;
  });
  if (oldItemOptions.length > 0) {
    const selectedValue = String(sheet.getRange(1, 2).getValue()).trim();
    const selectedItemId = getPlaidReplacementSelectedItemId(selectedValue);
    const matchingOption = oldItemOptions.find(option =>
      getPlaidReplacementSelectedItemId(option) === selectedItemId
    );
    sheet.getRange(1, 2)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(oldItemOptions, true)
        .setAllowInvalid(false)
        .build());
    if (matchingOption) sheet.getRange(1, 2).setValue(matchingOption);
  }
  sheet.getRange(4, 2).setValue(
    "Select the old bank/Item in B1. Review every decision in column A. Use the Replacement action dropdown in B5 to finalize or cancel."
  );
  return sheet;
}

function getPlaidReplacementSelectedItemId(selection) {
  const value = String(selection || "").trim();
  const delimiterIndex = value.lastIndexOf(" — ");
  return delimiterIndex >= 0 ? value.slice(delimiterIndex + 3).trim() : value;
}

function normalizeReplacementDate(dateValue) {
  return dateValue instanceof Date
    ? Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : String(dateValue);
}

function getReplacementDuplicateIndex() {
  const index = {};
  const add = (date, amount, description, txnId) => {
    const signature = [
      normalizeReplacementDate(date),
      String(Number(amount)),
      String(description || "").trim().toLowerCase()
    ].join("|");
    if (!index[signature]) index[signature] = [];
    if (!index[signature].includes(String(txnId))) index[signature].push(String(txnId));
  };

  const ledger = getFinanceSheet(LEDGER_SHEET);
  if (ledger.getLastRow() >= 2) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, 4).getValues().forEach(row => {
      add(row[1], row[2], row[3], row[0]);
    });
  }

  const unknown = getFinanceSheet(UNKNOWN_SHEET);
  if (unknown.getLastRow() >= 2) {
    unknown.getRange(2, 1, unknown.getLastRow() - 1, 6).getValues().forEach(row => {
      add(row[1], row[3], row[2], row[0]);
    });
  }
  return index;
}

function stagePlaidReplacementTransaction(sheet, txn, duplicateIndex, rowByTransactionId) {
  if (txn.date < SYNC_START_DATE) return;

  const transactionId = String(txn.transaction_id);
  const existingRow = rowByTransactionId[transactionId];
  const signature = [
    normalizeReplacementDate(txn.date),
    String(Number(txn.amount)),
    String(txn.name || txn.merchant_name || "").trim().toLowerCase()
  ].join("|");
  const possibleDuplicates = duplicateIndex[signature] || [];
  const automaticDecision = possibleDuplicates.length === 1
    ? `REPLACE ${possibleDuplicates[0]}`
    : "IMPORT";
  const values = [
    automaticDecision, possibleDuplicates.length > 1 ? "REVIEW" : "AUTO_DECISION",
    transactionId, txn.date, txn.amount,
    txn.name || txn.merchant_name || "", txn.merchant_name || "",
    txn.personal_finance_category?.primary || "",
    txn.personal_finance_category?.detailed || "",
    txn.account_id || "", possibleDuplicates.join(", ")
  ];

  if (existingRow) {
    const previous = sheet.getRange(existingRow, 1, 1, 2).getValues()[0];
    const previousDecision = String(previous[0] || "").trim();
    const previousStatus = String(previous[1] || "").trim().toUpperCase();
    if (
      previousDecision &&
      (previousStatus === "REVIEW" || previousStatus === "AUTO_DECISION")
    ) {
      values[0] = previousDecision;
      values[1] = previousStatus;
    } else if (
      previousStatus === "IMPORTED" ||
      previousStatus === "REPLACED" ||
      previousStatus === "SKIPPED" ||
      previousStatus === "REMOVED"
    ) {
      values[0] = previousDecision;
      values[1] = previousStatus;
    }
    sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
  } else {
    const row = sheet.getLastRow() + 1;
    ensureSheetHasRows(sheet, row);
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    rowByTransactionId[transactionId] = row;
  }
}

function stagePendingPlaidItemTransactions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return {
      status: "in_progress",
      item_id: financeScriptProperties.getProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID")
    };
  }

  try {
    return stagePendingPlaidItemTransactionsLocked();
  } finally {
    lock.releaseLock();
  }
}

function stagePendingPlaidItemTransactionsLocked() {
  const itemId = financeScriptProperties.getProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  if (!itemId) throw new Error("There is no newly linked Plaid Item waiting for replacement review.");

  const accessToken = financeScriptProperties.getProperty(`PLAID_ACCESS_TOKEN_${itemId}`);
  if (!accessToken) throw new Error(`Cannot stage Plaid Item ${itemId}: its access token is missing.`);

  const sheet = getPlaidReplacementReviewSheet(itemId);
  const lastRow = sheet.getLastRow();
  const rowByTransactionId = {};
  if (lastRow >= PLAID_REPLACEMENT_DATA_START_ROW) {
    sheet.getRange(PLAID_REPLACEMENT_DATA_START_ROW, 3, lastRow - PLAID_REPLACEMENT_DATA_START_ROW + 1, 1)
      .getValues()
      .forEach((row, index) => {
        rowByTransactionId[String(row[0])] = PLAID_REPLACEMENT_DATA_START_ROW + index;
      });
  }
  const duplicateIndex = getReplacementDuplicateIndex();
  const sessionStartCursor = financeScriptProperties.getProperty(`PLAID_CURSOR_${itemId}`) || "";
  let cursor = sessionStartCursor;
  let hasMore = true;
  const startTime = Date.now();
  const TIME_BUDGET_MS = 5 * 60 * 1000;

  sheet.getRange(3, 2).setValue("Staging in progress");
  while (hasMore) {
    if (Date.now() - startTime > TIME_BUDGET_MS) {
      sheet.getRange(3, 2).setValue("More pages remain; automatic staging will resume on the next inbox trigger");
      return { status: "in_progress", item_id: itemId };
    }

    const result = fetchNextSyncPage(accessToken, cursor);
    if (!result.success) {
      if (result.error === "TRANSACTIONS_SYNC_MUTATION_DURING_PAGINATION") {
        cursor = sessionStartCursor;
        continue;
      }
      throw new Error(`Could not stage transactions for Plaid Item ${itemId}: ${result.error}`);
    }

    const syncData = result.data;
    (syncData.added || []).forEach(txn =>
      stagePlaidReplacementTransaction(sheet, txn, duplicateIndex, rowByTransactionId)
    );
    (syncData.modified || []).forEach(txn =>
      stagePlaidReplacementTransaction(sheet, txn, duplicateIndex, rowByTransactionId)
    );
    (syncData.removed || []).forEach(removed => {
      const row = rowByTransactionId[String(removed.transaction_id)];
      if (row) sheet.getRange(row, 2).setValue("REMOVED");
    });

    cursor = syncData.next_cursor;
    financeScriptProperties.setProperty(`PLAID_CURSOR_${itemId}`, cursor);
    hasMore = syncData.has_more;
  }

  sheet.getRange(3, 2).setValue("Complete — review all rows before finalizing");
  return { status: "complete", item_id: itemId, staged_transactions: Math.max(0, sheet.getLastRow() - PLAID_REPLACEMENT_DATA_START_ROW + 1) };
}

function finalizePlaidItemReplacement() {
  const itemId = financeScriptProperties.getProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  if (!itemId) throw new Error("There is no staged Plaid Item replacement to finalize.");

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    financeScriptProperties.setProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID", itemId);
    return { status: "in_progress", item_id: itemId };
  }

  try {
    return finalizePlaidItemReplacementLocked();
  } catch (error) {
    financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID");
    financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID");
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function finalizePlaidItemReplacementLocked() {
  const newItemId = financeScriptProperties.getProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  if (!newItemId) throw new Error("There is no staged Plaid Item replacement to finalize.");

  const startTime = Date.now();
  const TIME_BUDGET_MS = 4 * 60 * 1000;
  financeScriptProperties.setProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID", newItemId);
  const newAccessToken = financeScriptProperties.getProperty(`PLAID_ACCESS_TOKEN_${newItemId}`);
  if (!newAccessToken) {
    throw new Error(`Cannot finalize Plaid Item ${newItemId}: its access token is missing.`);
  }
  const newCursor = financeScriptProperties.getProperty(`PLAID_CURSOR_${newItemId}`);
  if (newCursor === null) {
    throw new Error(`Cannot finalize Plaid Item ${newItemId}: its staged sync cursor is missing. Run stagePendingPlaidItemTransactions() again.`);
  }

  const sheet = getPlaidReplacementReviewSheet(newItemId);
  const stagingStatus = String(sheet.getRange(3, 2).getValue());
  if (
    stagingStatus.indexOf("Complete") !== 0 &&
    stagingStatus.indexOf("Finalization in progress") !== 0
  ) {
    throw new Error("The new Item's history is not fully staged. Run stagePendingPlaidItemTransactions() until it reports complete.");
  }

  const oldItemId = getPlaidReplacementSelectedItemId(sheet.getRange(1, 2).getValue());
  if (!oldItemId || oldItemId === newItemId) {
    throw new Error("Select the existing bank/Item to replace from the dropdown in cell B1 of the Plaid Review sheet.");
  }
  const itemIds = JSON.parse(financeScriptProperties.getProperty("PLAID_ITEM_IDS") || "[]");
  if (!itemIds.includes(oldItemId)) {
    throw new Error(`The old Item ID "${oldItemId}" is not in the active linked Item list.`);
  }
  if (itemIds.includes(newItemId)) {
    throw new Error(`The new Item ID "${newItemId}" is already active; refusing to replace an Item with itself.`);
  }

  const lastRow = sheet.getLastRow();
  const count = Math.max(0, lastRow - PLAID_REPLACEMENT_DATA_START_ROW + 1);
  const rows = count
    ? sheet.getRange(PLAID_REPLACEMENT_DATA_START_ROW, 1, count, 11).getValues()
    : [];
  for (let i = 0; i < rows.length; i++) {
    const decision = String(rows[i][0]).trim();
    const normalizedDecision = decision.toUpperCase();
    const status = String(rows[i][1]).trim().toUpperCase();
    if (status === "REMOVED" || status === "IMPORTED" || status === "REPLACED" || status === "SKIPPED") continue;

    const replaceMatch = decision.match(/^REPLACE\s+(.+)$/i);
    const suggestedIds = String(rows[i][10] || "").split(", ").filter(Boolean);
    if (
      normalizedDecision !== "IMPORT" &&
      normalizedDecision !== "SKIP" &&
      (!replaceMatch || !suggestedIds.includes(replaceMatch[1].trim()))
    ) {
      throw new Error(`Set Decision to IMPORT, SKIP, or REPLACE <suggested old transaction ID> for every staged transaction. First invalid or unresolved decision: ${PLAID_REPLACEMENT_DATA_START_ROW + i}.`);
    }
  }

  syncFixedPoolsFromAnnualOverview();

  let hasChanges = false;
  let processedRows = 0;
  for (let index = 0; index < rows.length; index++) {
    if (Date.now() - startTime >= TIME_BUDGET_MS) {
      sheet.getRange(3, 2).setValue(`Finalization in progress; ${processedRows} rows processed this run, automatic continuation is scheduled`);
      return { status: "in_progress", item_id: newItemId, processed_rows: processedRows };
    }

    const row = rows[index];
    const status = String(row[1]).trim().toUpperCase();
    if (status === "REMOVED" || status === "IMPORTED" || status === "REPLACED" || status === "SKIPPED") continue;
    const sheetRow = PLAID_REPLACEMENT_DATA_START_ROW + index;
    const decision = String(row[0]).trim();
    const normalizedDecision = decision.toUpperCase();
    if (normalizedDecision === "SKIP") {
      sheet.getRange(sheetRow, 2).setValue("SKIPPED");
      processedRows++;
      continue;
    }

    const txn = {
      transaction_id: String(row[2]),
      date: row[3] instanceof Date ? normalizeReplacementDate(row[3]) : String(row[3]),
      amount: Number(row[4]),
      name: String(row[5] || ""),
      merchant_name: String(row[6] || ""),
      personal_finance_category: {
        primary: String(row[7] || ""),
        detailed: String(row[8] || "")
      }
    };
    const replaceMatch = decision.match(/^REPLACE\s+(.+)$/i);
    if (replaceMatch) {
      const oldTransactionId = replaceMatch[1].trim();
      const removal = applyTransaction({ transaction_id: oldTransactionId }, true);
      if (removal.error) {
        throw new Error(`Could not remove old transaction ${oldTransactionId} before replacing it with ${txn.transaction_id}: ${removal.error}`);
      }
      if (removal.changed) hasChanges = true;
    }

    const outcome = applyTransaction(txn, false);
    if (outcome.error) {
      throw new Error(`Could not import staged transaction ${txn.transaction_id}: ${outcome.error}`);
    }
    if (outcome.changed) hasChanges = true;
    sheet.getRange(sheetRow, 2).setValue(replaceMatch ? "REPLACED" : "IMPORTED");
    processedRows++;
  }

  const hasPreviouslyAppliedTransactions = rows.some(row => {
    const status = String(row[1]).trim().toUpperCase();
    return status === "IMPORTED" || status === "REPLACED";
  });
  if (hasChanges || hasPreviouslyAppliedTransactions) {
    [LIFESTYLE_SHEET, MANDATORY_SHEET].forEach(sheetName => {
      recomputeAnnualRollup(getFinanceSheet(sheetName), sheetName);
    });
  }

  const updatedItemIds = itemIds.map(itemId => itemId === oldItemId ? newItemId : itemId);
  financeScriptProperties.setProperty("PLAID_ITEM_IDS", JSON.stringify(updatedItemIds));
  financeScriptProperties.deleteProperty(`PLAID_ACCESS_TOKEN_${oldItemId}`);
  financeScriptProperties.deleteProperty(`PLAID_CURSOR_${oldItemId}`);
  financeScriptProperties.deleteProperty(`PLAID_INSTITUTION_NAME_${oldItemId}`);
  financeScriptProperties.deleteProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_THREAD_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_STARTED_NOTIFIED_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_COMPLETE_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_LAST_ERROR");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_REVIEW_SHEET");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_REVIEW_SHEET_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID");
  let reviewSheetDeleted = false;
  let reviewSheetDeleteError = null;
  try {
    getFinanceSpreadsheet().deleteSheet(sheet);
    reviewSheetDeleted = true;
  } catch (error) {
    reviewSheetDeleteError = error.message || String(error);
    console.error(`Replacement completed, but review sheet "${sheet.getName()}" could not be deleted: ${reviewSheetDeleteError}`);
  }
  return {
    status: "success",
    old_item_id: oldItemId,
    new_item_id: newItemId,
    review_sheet_deleted: reviewSheetDeleted,
    review_sheet_delete_error: reviewSheetDeleteError
  };
}

function cancelPlaidItemReplacement() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error("Another replacement staging or finalization run is active. Wait for it to finish, then retry cancellation.");
  }

  try {
    return cancelPlaidItemReplacementLocked();
  } finally {
    lock.releaseLock();
  }
}

function cancelPlaidItemReplacementLocked() {
  const newItemId = financeScriptProperties.getProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  const replacementLinkInProgress =
    financeScriptProperties.getProperty("PLAID_LINK_MODE") === "REPLACEMENT_REVIEW";
  if (!newItemId && !replacementLinkInProgress) {
    throw new Error("There is no pending Plaid Item replacement or replacement Link session to cancel.");
  }

  let reviewSheet = null;
  if (newItemId) {
    const sheetName = getPlaidReplacementReviewSheetName(newItemId);
    reviewSheet = getFinanceSpreadsheet().getSheetByName(sheetName);
    if (reviewSheet) {
      if (String(reviewSheet.getRange(2, 2).getValue()) !== String(newItemId)) {
        throw new Error(`Refusing to delete "${sheetName}": it does not identify the pending Plaid Item ${newItemId}.`);
      }

      const lastRow = reviewSheet.getLastRow();
      if (lastRow >= PLAID_REPLACEMENT_DATA_START_ROW) {
        const statuses = reviewSheet.getRange(
          PLAID_REPLACEMENT_DATA_START_ROW,
          2,
          lastRow - PLAID_REPLACEMENT_DATA_START_ROW + 1,
          1
        ).getValues();
        if (statuses.some(row => {
          const status = String(row[0]).trim().toUpperCase();
          return status === "IMPORTED" || status === "REPLACED";
        })) {
          throw new Error("This replacement has imported transactions. The review sheet cannot be deleted; finalize the replacement or reconcile imported transactions first.");
        }
      }
    }
  }

  if (newItemId) {
    financeScriptProperties.deleteProperty(`PLAID_ACCESS_TOKEN_${newItemId}`);
    financeScriptProperties.deleteProperty(`PLAID_CURSOR_${newItemId}`);
    financeScriptProperties.deleteProperty(`PLAID_INSTITUTION_NAME_${newItemId}`);
  }
  financeScriptProperties.deleteProperty("PLAID_PENDING_REPLACEMENT_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_THREAD_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_STARTED_NOTIFIED_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_COMPLETE_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_LINK_STARTED_AT");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_LINK_LAST_ERROR");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_STAGING_LAST_ERROR");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_REVIEW_SHEET");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_REVIEW_SHEET_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_FINALIZATION_NOTIFIED_ITEM_ID");
  financeScriptProperties.deleteProperty("PLAID_LINK_MODE");
  CacheService.getScriptCache().remove("CURRENT_LINK_TOKEN");
  if (reviewSheet) getFinanceSpreadsheet().deleteSheet(reviewSheet);
  return {
    status: "cancelled",
    item_id: newItemId || null,
    review_sheet_deleted: !!reviewSheet
  };
}

// =========================================================================
// RE-SYNC FROM THE START — clears only Plaid cursors for the active linked
// items, then replays their available history through the normal sync path.
// Existing spreadsheet data is left untouched.
// =========================================================================
function resyncTransactionsFromStart() {
  const itemIds = JSON.parse(financeScriptProperties.getProperty('PLAID_ITEM_IDS') || "[]");
  if (itemIds.length === 0) {
    throw new Error("Cannot re-sync transactions: no linked Plaid items were found.");
  }

  const missingTokens = itemIds.filter(itemId =>
    !financeScriptProperties.getProperty(`PLAID_ACCESS_TOKEN_${itemId}`)
  );
  if (missingTokens.length > 0) {
    throw new Error(`Cannot re-sync transactions: missing access token(s) for Plaid item(s): ${missingTokens.join(", ")}.`);
  }

  itemIds.forEach(itemId => {
    financeScriptProperties.deleteProperty(`PLAID_CURSOR_${itemId}`);
  });

  console.log(`Plaid cursors cleared for ${itemIds.length} linked item(s). Starting sync from the beginning of each item's available history; spreadsheet data was not cleared.`);
  return runDailyFinanceSync();
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
  financeScriptProperties.setProperty('FINANCE_SPREADSHEET_NAME', newSpreadsheetName);

  const itemIds = JSON.parse(financeScriptProperties.getProperty('PLAID_ITEM_IDS') || "[]");
  itemIds.forEach(itemId => {
    financeScriptProperties.deleteProperty(`PLAID_CURSOR_${itemId}`);
  });

  console.log(`Finance target switched to "${newSpreadsheetName}". Cursors cleared for ${itemIds.length} item(s) — the next sync will pull full available history, filtered to ${SYNC_START_DATE} onward.`);
}

// =========================================================================
// ONE-TIME SETUP — creates the new sheets this pipeline depends on, with
// headers, if they don't already exist. Run this once manually.
// =========================================================================
function setupFinanceSheets() {
  const ss = getFinanceSpreadsheet();
  ensureMonthlyFinanceSheets(ss);
  removeCategoryReplyTriggers();
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_THREAD_ID");
  financeScriptProperties.deleteProperty("PLAID_REPLACEMENT_APPROVAL_ITEM_ID");

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
    sheet.getRange(1, 1, 1, 11).setValues([[
      "Transaction ID", "Date", "Amount", "Description",
      "Target Sheet", "Target Field", "Month", "Year",
      "Override Category", "Default Target Sheet", "Default Target Field"
    ]]);
  }
  const ledgerSheet = ss.getSheetByName(LEDGER_SHEET);
  if (ledgerSheet.getMaxColumns() < 11) {
    ledgerSheet.insertColumnsAfter(ledgerSheet.getMaxColumns(), 11 - ledgerSheet.getMaxColumns());
  }
  ledgerSheet.getRange(1, 9, 1, 3).setValues([[
    "Override Category", "Default Target Sheet", "Default Target Field"
  ]]);
  if (
    ledgerSheet.getMaxRows() > 1 &&
    ss.getSheetByName(LIFESTYLE_SHEET) &&
    ss.getSheetByName(MANDATORY_SHEET)
  ) {
    ledgerSheet.getRange(2, LEDGER_OVERRIDE_COLUMN, ledgerSheet.getMaxRows() - 1, 1)
      .setDataValidation(getTransactionOverrideValidationRule());
  }
  ensureTransactionLedgerEditTrigger(ss);

  if (!ss.getSheetByName(UNKNOWN_SHEET)) {
    const sheet = ss.insertSheet(UNKNOWN_SHEET);
    sheet.getRange(1, 1, 1, 9).setValues([[
      "Transaction ID", "Date", "Description", "Amount", "Plaid Category", "Merchant Name",
      "Plaid Category (primary)", "Plaid Category (detailed)", "Assign to Category"
    ]]);
  }
  const unknownSheet = ss.getSheetByName(UNKNOWN_SHEET);
  if (unknownSheet.getMaxColumns() < UNKNOWN_CATEGORY_COLUMN) {
    unknownSheet.insertColumnsAfter(
      unknownSheet.getMaxColumns(),
      UNKNOWN_CATEGORY_COLUMN - unknownSheet.getMaxColumns()
    );
  }
  unknownSheet.getRange(1, UNKNOWN_CATEGORY_COLUMN).setValue("Assign to Category");
  setUnknownCategoryValidation(
    unknownSheet,
    2,
    Math.max(unknownSheet.getLastRow() - 1, 1)
  );

  if (ss.getSheetByName(ANNUAL_OVERVIEW_SHEET)) {
    ensureMonthSectionsUpToDate();
  }

  console.log("Finance sheet setup complete. Monthly sheets and current-year month sections are ready; annual budgets also create the fixed pools. Transaction Ledger overrides, Uncategorized Transactions dropdowns, and the spreadsheet edit trigger are ready.");
}
