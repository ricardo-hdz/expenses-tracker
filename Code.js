/**
 * Script to process personal expenses and update Master budget report.
 */

const EXPENSE_TYPES = [
  'Auto',
  'Books',
  'Electronics',
  'Entertainment',
  'Food',
  'Home',
  'Household',
  'Income',
  'Misc',
  'Personal',
  'Pet',
  'Transportation',
  'Travel',
  'Utilities',
  'Vacation',
  'Weekend',
  'iTunes',
  'Claveles'
];

const EXPENSE_SUBTYPES = {
  'Utilities': [
    'Gas',
    'Water',
    'Electricity',
    'Cellphone',
    'Internet'
  ],
  'Misc': [
    'Cerveza'
  ],
  'Personal': [
    'Clothing',
    'Gym'
  ],
  'Weekend': [
    'Pistos'
  ],
  'Household': [
    'weekend home'
  ]
};

const MONTH_COLUMN_MAP = {
  'Jan': 2,
  'Feb': 3,
  'Mar': 4,
  'Apr': 5,
  'May': 6,
  'Jun': 7,
  'Jul': 8,
  'Aug': 9,
  'Sep': 10,
  'Oct': 11,
  'Nov': 12,
  'Dec': 13
};

// Master budget column in master sheet (Column N = 14)
const MASTER_BUDGET_COLUMN = 14;

// Month Chart Coordinates
const SUM_CHART_START_ROW = 2;
const SUM_CHART_START_COL = 10; // Column J

// Data ranges in month sheet for entries
const CATEGORY_RANGE_A1 = "C2:C200";
const SUBCATEGORY_RANGE_A1 = "D2:D200";
const AMOUNT_RANGE_A1 = "H2:H200";

// ==========================================
// Application Menu
// ==========================================

function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const menuEntries = [
    { name: "Process Month", functionName: "processMonth" },
    { name: "Process All Months", functionName: "processAllMonths" }
  ];
  spreadsheet.addMenu("ExpenseTracker", menuEntries);
}

// ==========================================
// Main Workflow Functions
// ==========================================

/**
 * Main process function for currently active month sheet.
 */
function processMonth() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const month = sheet.getName();

  if (!MONTH_COLUMN_MAP[month]) {
    SpreadsheetApp.getUi().alert(
      `Active sheet '${month}' is not a valid month sheet (Jan - Dec). Please select a month tab before running.`
    );
    return;
  }

  const masterSheet = spreadsheet.getSheetByName('Master');
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Master' not found in this spreadsheet.");
    return;
  }

  const numRows = setupSumChart(sheet);
  SpreadsheetApp.flush(); // Ensure formulas calculate before reading values

  copyMonthTotalsToMaster(sheet, masterSheet, month, numRows);
  verifyMonthBudget(masterSheet, month, numRows);

  spreadsheet.toast(`Processed ${month} successfully.`, 'Expense Tracker', 5);
}

/**
 * Processes all 12 month sheets (Jan through Dec) sequentially.
 */
function processAllMonths() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName('Master');
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Master' not found in this spreadsheet.");
    return;
  }

  let processedCount = 0;
  for (const month of Object.keys(MONTH_COLUMN_MAP)) {
    const sheet = spreadsheet.getSheetByName(month);
    if (sheet) {
      const numRows = setupSumChart(sheet);
      SpreadsheetApp.flush();
      copyMonthTotalsToMaster(sheet, masterSheet, month, numRows);
      verifyMonthBudget(masterSheet, month, numRows);
      processedCount++;
    }
  }

  spreadsheet.toast(`Processed ${processedCount} months successfully.`, 'Expense Tracker', 5);
}

// ==========================================
// Chart & Formula Setup
// ==========================================

/**
 * Dynamically constructs and writes the monthly sum chart (titles and SUMIF formulas)
 * in columns J, K, and L using batch operations.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number} Number of rows in the summary chart.
 */
function setupSumChart(sheet) {
  const targetSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { chartValues, chartFormulas, totalRows } = generateChartData(SUM_CHART_START_ROW);

  // Clear previous chart range in columns J, K, L (cols 10-12)
  const existingLastRow = targetSheet.getLastRow();
  const rowsToClear = Math.max(
    existingLastRow >= SUM_CHART_START_ROW ? existingLastRow - SUM_CHART_START_ROW + 1 : 0,
    totalRows
  );
  if (rowsToClear > 0) {
    targetSheet.getRange(SUM_CHART_START_ROW, SUM_CHART_START_COL, rowsToClear, 3).clearContent();
  }

  // Batch write categories and subcategories (Columns J and K)
  targetSheet.getRange(SUM_CHART_START_ROW, SUM_CHART_START_COL, chartValues.length, 2).setValues(chartValues);

  // Batch write formulas (Column L)
  targetSheet.getRange(SUM_CHART_START_ROW, SUM_CHART_START_COL + 2, chartFormulas.length, 1).setFormulas(chartFormulas);

  return totalRows;
}

/**
 * Generates the in-memory 2D arrays for chart values and formulas.
 * @param {number} startRow
 * @returns {{ chartValues: Array<Array>, chartFormulas: Array<Array>, totalRows: number }}
 */
function generateChartData(startRow) {
  const chartValues = [];
  const chartFormulas = [];
  let currentRow = startRow;

  for (const category of EXPENSE_TYPES) {
    chartValues.push([category, '']);
    chartFormulas.push([`=SUMIF(${CATEGORY_RANGE_A1}, J${currentRow}, ${AMOUNT_RANGE_A1}) * -1`]);
    currentRow++;

    if (EXPENSE_SUBTYPES[category] && EXPENSE_SUBTYPES[category].length > 0) {
      for (const subCategory of EXPENSE_SUBTYPES[category]) {
        chartValues.push(['', subCategory]);
        chartFormulas.push([`=SUMIF(${SUBCATEGORY_RANGE_A1}, K${currentRow}, ${AMOUNT_RANGE_A1}) * -1`]);
        currentRow++;
      }
    }
  }

  const lastCategoryRow = currentRow - 1;
  // Total row: Sums only main category rows (where column J is not empty) to avoid double-counting subcategories
  chartValues.push(['Total', '']);
  chartFormulas.push([`=SUMIF(J${startRow}:J${lastCategoryRow}, "<>", L${startRow}:L${lastCategoryRow})`]);

  return {
    chartValues,
    chartFormulas,
    totalRows: chartValues.length
  };
}

/**
 * Returns the total number of rows generated in the chart.
 * @returns {number}
 */
function getChartRowCount() {
  let count = 0;
  for (const category of EXPENSE_TYPES) {
    count++;
    if (EXPENSE_SUBTYPES[category] && EXPENSE_SUBTYPES[category].length > 0) {
      count += EXPENSE_SUBTYPES[category].length;
    }
  }
  count++; // Total row
  return count;
}

// ==========================================
// Master Sheet Synchronization & Verification
// ==========================================

/**
 * Copies the month totals (Column L) to the Master spreadsheet in a single batch.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} monthSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet
 * @param {string} month
 * @param {number} [numRows]
 */
function copyMonthTotalsToMaster(monthSheet, masterSheet, month, numRows) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = monthSheet || spreadsheet.getActiveSheet();
  const targetMaster = masterSheet || spreadsheet.getSheetByName('Master');
  const monthName = month || sourceSheet.getName();

  const monthColumn = MONTH_COLUMN_MAP[monthName];
  if (!monthColumn || !targetMaster) return;

  const rowCount = numRows || getChartRowCount();
  const columnL = SUM_CHART_START_COL + 2; // Column 12 (L)

  const values = sourceSheet.getRange(SUM_CHART_START_ROW, columnL, rowCount, 1).getValues();
  targetMaster.getRange(SUM_CHART_START_ROW, monthColumn, rowCount, 1).setValues(values);
}

/**
 * Compares actual expenses against the budget in the Master sheet in a single batch
 * and color-codes over-budget rows in red with alternating background colors.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet
 * @param {string} month
 * @param {number} [numRows]
 */
function verifyMonthBudget(masterSheet, month, numRows) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetMaster = masterSheet || spreadsheet.getSheetByName('Master');
  const monthName = month || spreadsheet.getActiveSheet().getName();

  const monthColumn = MONTH_COLUMN_MAP[monthName];
  if (!monthColumn || !targetMaster) return;

  const rowCount = numRows || getChartRowCount();

  const expenseValues = targetMaster.getRange(SUM_CHART_START_ROW, monthColumn, rowCount, 1).getValues();
  const budgetValues = targetMaster.getRange(SUM_CHART_START_ROW, MASTER_BUDGET_COLUMN, rowCount, 1).getValues();

  const backgrounds = [];
  for (let i = 0; i < rowCount; i++) {
    const rowNumber = SUM_CHART_START_ROW + i;
    const expense = Number(expenseValues[i][0]) || 0;
    const budget = Number(budgetValues[i][0]) || 0;

    if (budget > 0 && expense > budget) {
      backgrounds.push(['red']);
    } else if (budget === 0 && expense > 0) {
      backgrounds.push(['red']);
    } else {
      if (rowNumber % 2 !== 0) {
        backgrounds.push(['#a4c2f4']); // RGB(164, 194, 244)
      } else {
        backgrounds.push(['#ffffff']); // White
      }
    }
  }

  targetMaster.getRange(SUM_CHART_START_ROW, monthColumn, rowCount, 1).setBackgrounds(backgrounds);
}

// ==========================================
// Backwards Compatibility Stubs
// ==========================================

function setCategoryMonthlyTotals(row, column, conceptName, rangeOfEntries) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, column).setFormula(`=SUMIF(${rangeOfEntries}, ${conceptName}, ${AMOUNT_RANGE_A1}) * -1`);
}

function setMonthTotals(row, column, monthTotal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, column).setValue(monthTotal);
}