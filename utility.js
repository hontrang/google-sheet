function logTime(sheet, cell) {
  SpreadsheetApp.getActive()
    .getSheetByName(sheet)
    .getRange(cell)
    .setValue(new Date());
}

function getDate(number) {
  let date = new Date(number);
  return (
    date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate()
  );
}

function ganDuLieuVaoCot(cotThamChieu, cotDoiDuLieu) {
  let sheetDuLieu = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  sheetDuLieu.getRange(cotDoiDuLieu + ":" + cotDoiDuLieu).clearContent();
  let data = sheetDuLieu.getRange(cotThamChieu + "2").getValue();
  Logger.log(data);
}

function ghiDuLieuVaoDay(data, sheetName, row, column, width, height) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  // xoá trắng range
  sheet.getRange(row, column, width, height).clearContent();

  sheet.getRange(row, column, width, height).setValues(data);
}

function ghiDuLieuVaoO(data, sheetName, row, column, width, height) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  // xoá trắng range
  sheet.getRange(row, column, width, height).clearContent();

  sheet.getRange(row, column, width, height).setValue(data);
}

function layDuLieuTrongO(sheetName, row, column) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  return sheet.getRange(row, column).getValue();
}
