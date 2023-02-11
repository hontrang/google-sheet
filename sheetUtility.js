var SheetUtility = {
  ghiDuLieuVaoDay: function (data, sheetName, row, column) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // xoá trắng range
    sheet.getRange(row, column, data.length, data[0].length).clearContent();

    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  },
  ghiDuLieuVaoO: function (data, sheetName, row, column) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // xoá trắng range
    sheet.getRange(row, column, 1, 1).clearContent();

    sheet.getRange(row, column, 1, 1).setValue(data);
  },
  layDuLieuTrongO: function (sheetName, row, column) {
    return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(row, column).getValue();
  },
  layGiaTriTheoCot: function (activeSheet, rowIndex, columnIndex) {
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
}

