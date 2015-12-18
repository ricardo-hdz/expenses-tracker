/**
 * Script to process personal expenses using iXpenseit.
 *
 * Feature backlog:
 * - Remove business checking entries
 * - Include a "Process all months" option
 *
 */

var EXPENSE_TYPES = [
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
  'iTunes'
];
; Sign to use for amount values
; Some expense apps use minus (-) for expenses, others use positive values
var SIGN_MULTIPLIER = 1;
var EXPENSE_SUBTYPES = [];
EXPENSE_SUBTYPES['Utilities'] = [
  'Gas',
  'Water',
  'Electricity',
  'Cellphone',
  'Internet'
];
EXPENSE_SUBTYPES['Misc'] = [
  'Cerveza'
];
EXPENSE_SUBTYPES['Personal'] = [
  'Clothing',
  'Gym'
];
EXPENSE_SUBTYPES['Weekend'] = [
  'Pistos'
];

var MONTH_COLUMN_MAP = {
  'Jan': '2',
  'Feb': '3',
  'Mar': '4',
  'Apr': '5',
  'May': '6',
  'Jun': '7',
  'Jul': '8',
  'Aug': '9',
  'Sep': '10',
  'Oct': '11',
  'Nov': '12',
  'Dec': '13'
};

// Master budget column in master sheet
var MASTER_BUDGET_COLUMN = 14;

// Month Chart
var SUM_CHART_RANGE = 'J2';
var SUM_TOTAL_RANGE = 'L2';

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var masterSheet = spreadsheet.getSheetByName('Master');
var sheet = spreadsheet.getActiveSheet();

// Category, Subcategory and amount ranges (from current month)
var CATEGORY_RANGE = sheet.getRange("C2:C200");
var SUBCATEGORY_RANGE = sheet.getRange("D2:D200");
var AMOUNT_RANGE = sheet.getRange("H2:H200");
var MONTHLY_SUM = sheet.getRange("L2:L28");

function onOpen() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  masterSheet = spreadsheet.getSheetByName('Master');
  sheet = spreadsheet.getActiveSheet();
  var menuEntries = [
    {name: "Process Month", functionName: "processMonth"}
  ];
  spreadsheet.addMenu("ExpenseTracker", menuEntries);
}

/**
 * Sets the sum chart (titles and formulas)
 *
 */
function setupSumChart() {
  var catRange,
      subRange,
      row,
      column,
      subtypeColumn,
      monthTotal = 0;

  catRange = sheet.getRange(SUM_CHART_RANGE);
  row = catRange.getRow();
  column = catRange.getColumn();

  for (var i = 0, category; (category = EXPENSE_TYPES[i]); i++) {
    catRange = sheet.getRange(row, column);
    catRange.setValue(category);
    // catRange.getA1Notations() returns the name of the category
    setCategoryMonthlyTotals(row, column + 2, catRange.getA1Notation(), CATEGORY_RANGE.getA1Notation());

    // Sume the value of the category to the total
    monthTotal = monthTotal + sheet.getRange(row, column + 2).getValue();

    //Display Subcategories, if any
    if (EXPENSE_SUBTYPES[category] && EXPENSE_SUBTYPES[category].length > 0) {
      subtypeColumn = column + 1;
      row++;
      for (var j = 0, subCategory; (subCategory = EXPENSE_SUBTYPES[category][j]); j++) {
        subRange = sheet.getRange(row, subtypeColumn);
        subRange.setValue(subCategory);
        // subRange.getA1Notations() returns the name of the category
        setCategoryMonthlyTotals(row, subtypeColumn + 1, subRange.getA1Notation(), SUBCATEGORY_RANGE.getA1Notation());
        row++;
      }
    } else {
      row++;
    }
  }
  setMonthTotals(row, column + 2, monthTotal);
}

/**
 * Sets the formula for the monthly totals (by category)
 */
function setCategoryMonthlyTotals(row, column, conceptName, rangeOfEntries) {
  sheet.getRange(row, column).setFormula("=SUMIF(" + rangeOfEntries + "," + conceptName + "," + AMOUNT_RANGE.getA1Notation() + ") * SIGN_MULTIPLIER");
}

function setMonthTotals(row, column, monthTotal) {
  spreadsheet.getActiveSheet().getRange(row, column).setValue(monthTotal);
}

/**
 * Copy the month totals to the master spreadsheet
 */
function copyMonthTotalsToMaster() {
  var column,
      month,
      range,
      row,
      sheet,
      values,
      spreadsheet,
      sheet;

  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getActiveSheet();

  month = ScriptProperties.getProperty('month');
  column = MONTH_COLUMN_MAP[month];
  values = MONTHLY_SUM.getValues();
  sheet = spreadsheet.setActiveSheet(masterSheet);
  row = 2;
  range = sheet.getRange(row, column);

  for (var i = 0, val; (val = values[i]); i++) {
    sheet.getRange(row, column).setValue(val);
    row++;
  }
}

/**
* Main process function.
**/
function processMonth() {
  var currentMonth;

  currentMonth = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  ScriptProperties.setProperty('month', currentMonth);

  setupSumChart();
  copyMonthTotalsToMaster();
  verifyMonthBudget();
}

function verifyMonthBudget() {
  var monthColumn,
      budgetAmount,
      expenseAmount,
      expenseCell,
      month,
      row,
      spreadsheet,
      sheet;

  month = ScriptProperties.getProperty('month');
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getActiveSheet();

  monthColumn = MONTH_COLUMN_MAP[month];

  for (row = 2; row < 26; row++) {
    expenseCell = sheet.getRange(row, monthColumn);
    expenseAmount = expenseCell.getValue();
    budgetAmount = sheet.getRange(row, MASTER_BUDGET_COLUMN).getValue();
    if (expenseAmount > budgetAmount) {
      expenseCell.setBackground('red');
    } else {
      if (row % 2) {
        expenseCell.setBackgroundRGB(164, 194, 244);
      } else {
        expenseCell.setBackgroundRGB(255, 255, 255);
      }
    }
  }
}