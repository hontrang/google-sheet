var SheetUtility = {
  ghiDuLieuVaoDay: function (data, sheetName, row, column) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    sheet.getRange(row, column, data.length, data[0].length).clearContent();
    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  },
  ghiDuLieuVaoDayTheoTen: function (data, sheetName, rowName, columnName) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

    let rowIndex = parseInt(rowName, 10) - 1;
    let columnIndex = this.columnToIndex(columnName) - 1;

    sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
    sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
  },
  ghiDuLieuVaoO: function (data, sheetName, cell) {
    SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).setValue(data);
  },
  layDuLieuTrongO: function (sheetName, cell) {
    return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).getValue();
  },
  layGiaTriTheoCot: function (activeSheet, rowIndex, columnIndex) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(activeSheet);
    range = sheet.getRange(rowIndex, columnIndex, sheet.getLastRow() - rowIndex + 1);
    // xoá phần tử rỗng trong mảng
    return range.getValues().filter(function (el) {
      return el != "";
    });
  },
  layDuLieuTrongCot: function (sheetName, column) {
    let values = SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(`${column}:${column}`).getValues();
    values = values.filter(String);
    // xoa giá trị trong ô title
    values.reverse().pop();
    return values;
  },
  columnToIndex: function (columnName) {
    let index = 0;
    let length = columnName.length;
    for (let i = 0; i < length; i++) {
      let charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += (charCode * Math.pow(26, length - i - 1));
    }
    return index;
  }
}

