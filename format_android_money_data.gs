var HEADING_COLUMN = {
  '6': '1', //Date
  '15': '2', //Type (default to personal)
  '4': '3', //Category
  '5': '4', //Subcategory
  '12': '5', //Vendor
  '7': '6', //Payment (check if empty)
  '2': '7', //Currency
  '3': '8', //Amount
  '9': '9' //Note
};

var MONTHS = [
  'Test',
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec'
];


var TRACKER_ID = '1dzNLLJ_gVZMcsS-Q5FtxHv9ka9SAKbsHXtJc6zAIHeA';

var MAX_ROW_NUMBER = 200;

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var formattedSheet = spreadsheet.getSheetByName('Formatted');
var sheet = spreadsheet.getActiveSheet();

function onOpen() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  formattedSheet = spreadsheet.getSheetByName('Formatted');
  sheet = spreadsheet.getActiveSheet();
  var menuEntries = [
    {name: "Format & Transfer Data", functionName: "transformData"}
  ];
  spreadsheet.addMenu("Android Menu", menuEntries);
}

/*
* Assigns a positive or negative value to the amount according to
* type of transaction: income or expense
*/
function transformAmount() {
  var columnExpenseType = 'G';
  var columnAmount = 'C';
  sheet = spreadsheet.getActiveSheet();

  var id;
  var cellValue;

  for (var i = 2; i < MAX_ROW_NUMBER; i++) {
    id = columnExpenseType + i;
        cell = sheet.getRange(id);
    cellValue = cell.getValue();
    id = columnAmount + i;
      
    if (cellValue != '') {
            cell = sheet.getRange(id);
            cell.setValue(0 - cell.getValue());
    } else {
      cell = sheet.getRange(id);
            cell.setValue(0 + cell.getValue());
    }
  }
}

/*
 * Copy the raw data to formatted spreadsheet
 */
function transformData() {
  transformAmount();
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getActiveSheet();

  for (var columnSource in HEADING_COLUMN) {
    var values = sheet.getRange(3, columnSource, MAX_ROW_NUMBER); //getRange(row, column, numRows)
    var targetColumn = HEADING_COLUMN[columnSource];
    values.copyValuesToRange(formattedSheet, targetColumn, targetColumn, 2, MAX_ROW_NUMBER);
  } 
}

/*
* Copies all data to tracker spreadsheet
*/
function copyDataToTracker() {
  // Get source data
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  formattedSheet = spreadsheet.getSheetByName('Formatted');
  var sourceData = formattedSheet.getDataRange().getValues()
  
  var targetSheetName = getCurrentMonth();
  
  // Copy data to tracker
  var trackerSpreadsheet = SpreadsheetApp.openById(TRACKER_ID).getSheetByName(targetSheetName);
  var targetRangeTop = trackerSpreadsheet.getLastRow();
  trackerSpreadsheet.getRange(1,1, sourceData.length, sourceData[0].length).setValues(sourceData);
}

/*
* Get the current month id for the expenses
* @returns {string}
*/
function getCurrentMonth() {
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = spreadsheet.getSheetByName('Formatted');

  var cell = sheet.getRange('A2');
  var dateString = cell.getValue() + '';
  
  var year        = dateString.substring(0,4);
  var month       = dateString.substring(4,6);
  var day         = dateString.substring(6,8);
  
  return MONTHS[month];
}