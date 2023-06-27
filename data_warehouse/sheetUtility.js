var SheetUtility = {
  SHEET_THAM_CHIEU: "tham chiếu",
  SHEET_BANG_THONG_TIN: "bảng thông tin",
  SHEET_DU_LIEU: "dữ liệu",
  SHEET_CHI_TIET_MA: "chi tiết mã",
  SHEET_CAU_HINH: "cấu hình",
  SHEET_DEBUG: "debug",
  KICH_THUOC_MANG_PHU: 20,
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
    try {
      sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      SheetLog.logDebug(data);
      console.error(e);
    }
  },
  ghiDuLieuVaoO: function (data, sheetName, cell) {
    SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).setValue(data);
  },
  layDuLieuTrongO: function (sheetName, cell) {
    return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).getValue();
  },
  layGiaTriTheoCot: function (activeSheet, rowIndex, columnIndex) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(activeSheet);
    const range = sheet.getRange(rowIndex, columnIndex, sheet.getLastRow() - rowIndex + 1);
    // xoá phần tử rỗng trong mảng
    return range.getValues().filter(function (el) {
      return el != "";
    });
  },
  layDuLieuTrongCot: function (sheetName, column) {
   const columnData = SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(`${column}:${column}`).getValues();

    const dataArray = [];
    for (const element of columnData) {
      const value = element[0];
      if (value !== "") {
        dataArray.push(value);
      }
    }
    // Xóa phần tử đầu tiên ("title") trong mảng dataArray
    dataArray.shift();
    return dataArray;
  },
  columnToIndex: function (columnName) {
    let index = 0;
    let length = columnName.length;
    for (let i = 0; i < length; i++) {
      let charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += (charCode * Math.pow(26, length - i - 1));
    }
    return index;
  },
  chen1HangVaoDauSheet: function (sheetName) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // Chèn một hàng mới vào vị trí đầu tiên của bảng
    sheet.insertRowsBefore(1, 1);

    // Trả về vị trí của hàng mới
    return 1;
  }
}

