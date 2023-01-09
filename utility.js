function logTime(sheet, cell) {
  SpreadsheetApp.getActive().getSheetByName(sheet).getRange(cell).setValue(new Date());
}

function getDate(number) {
  let date = new Date(number);
  return date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate();
}
