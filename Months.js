/**
 * Copy the month totals to the master spreadsheet
 */
function copyMonthTotalsExtra() {
  var column,
      month,
      range,
      row,
      sheet,
      values;
  
  month = spreadsheet.getActiveSheet().getName();
  column = MONTH_COLUMN_MAP[month];
  sheet = spreadsheet.setActiveSheet(masterSheet);
  values = sheet.getRange("L2:L21").getValues();
  range = sheet.getRange(row, column);
  row = 2;
  
  for (var i = 0, val; (val = values[i]); i++) {
    sheet.getRange(row, column).setValue(val);
    row++;
  }
}