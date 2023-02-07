function logTime(sheet, cell) {
  SpreadsheetApp.getActive()
    .getSheetByName(sheet)
    .getRange(cell)
    .setValue(moment().format("YYYY/MM/DD HH:mm:ss"));
}

function logStart(sheet, cell) {
  SpreadsheetApp.getActive()
    .getSheetByName(sheet)
    .getRange(cell)
    .setValue("...");
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

function ghiDuLieuVaoDay(data, sheetName, row, column) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  // xoá trắng range
  sheet.getRange(row, column, data.length, data[0].length).clearContent();

  sheet.getRange(row, column, data.length, data[0].length).setValues(data);
}

function ghiDuLieuVaoDayXoaCot(data, sheetName, row, column) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  // xoá trắng range
  sheet.getRange(row, column, sheet.getLastRow(), data[0].length).clearContent();

  sheet.getRange(row, column, data.length, data[0].length).setValues(data);
}

function ghiDuLieuVaoO(data, sheetName, row, column) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  // xoá trắng range
  sheet.getRange(row, column, 1, 1).clearContent();

  sheet.getRange(row, column, 1, 1).setValue(data);
}

function layDuLieuTrongO(sheetName, row, column) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  return sheet.getRange(row, column).getValue();
}

function layGiaTriTheoCot(activeSheet, rowIndex, columnIndex) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(activeSheet);
  range = sheet.getRange(
    rowIndex,
    columnIndex,
    sheet.getLastRow() - rowIndex + 1
  );
  // xoá phần tử rỗng trong mảng
  return range.getValues().filter(function (el) {
    return el != "";
  });
}
