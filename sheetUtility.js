var SheetUtility = {
  SHEET_THAM_CHIEU: "tham chiếu",
  SHEET_BANG_THONG_TIN: "bảng thông tin",
  SHEET_DU_LIEU: "dữ liệu",
  SHEET_CHI_TIET_MA: "chi tiết mã",
  SHEET_CAU_HINH: "cấu hình",
  SHEET_DEBUG: "debug",
  SHEET_GIA: "Giá",
  SHEET_KHOI_LUONG: "Khối Lượng",
  SHEET_KHOI_NGOAI_MUA: "KN Mua",
  SHEET_KHOI_NGOAI_BAN: "KN Bán",
  SHEET_HOSE: "HOSE",
  KICH_THUOC_MANG_PHU: 20,
  ghiDuLieuVaoDay: function (data, sheetName, row, column) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    sheet.getRange(row, column, data.length, data[0].length).clearContent();
    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  },
  ghiDuLieuVaoDayTheoTen: function (data, sheetName, rowName, columnName) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

    const rowIndex = parseInt(rowName, 10) - 1;
    const columnIndex = this.doiTenCotThanhChiSo(columnName) - 1;

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
  layDuLieuTrongOTheoTen: function (sheetName, cell) {
    return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).getValue();
  },
  layDuLieuTrongO: function (sheetName, row, col) {
    return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(row, col).getValue();
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
    return dataArray;
  },
  laySoHangTrongSheet: function (sheetName) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    return sheet.getLastRow();
  },
  doiTenCotThanhChiSo: function (columnName) {
    let index = 0;
    let length = columnName.length;
    for (let i = 0; i < length; i++) {
      let charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += (charCode * Math.pow(26, length - i - 1));
    }
    return index;
  },
  chen1HangVaoDauSheet: function (sheetName) {
    const index = 1;
    return this.chen1HangVaoSheet(sheetName, index);
  },
  chen1HangVaoSheet: function (sheetName, index) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    // Chèn một hàng mới vào vị trí đầu tiên của bảng
    sheet.insertRowsBefore(index, 1);

    // Trả về vị trí của hàng mới
    return 1;
  },
  layDuLieuTrongHang: function (sheetName, rowIndex) {
    let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

    // Lấy số lượng cột trong Sheet
    const numColumns = sheet.getLastColumn();

    // Lấy dữ liệu từ hàng
    const range = sheet.getRange(rowIndex, 1, 1, numColumns);
    const rowData = range.getValues();

    // rowData là một mảng 2 chiều, chúng ta cần phải lấy phần tử đầu tiên để có mảng 1 chiều
    return rowData[0];
  },
  ghiDuLieuVaoDayTheoTenThamChieu: function (data, sheetName, cotBatDau, cotThamChieu, giaTriThamChieu) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const hangBatDau = this.timHangCoGiaTriTrongCot(sheetName, cotThamChieu, giaTriThamChieu);
    const rowIndex = parseInt(hangBatDau, 10) - 1;
    const columnIndex = this.doiTenCotThanhChiSo(cotBatDau) - 1;

    sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
    try {
      sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      SheetLog.logDebug(data);
      console.error(e);
    }
  },
  timHangCoGiaTriTrongCot: function (sheetName, cotThamChieu, giaTriThamChieu) {
    const columnIndex = this.doiTenCotThanhChiSo(cotThamChieu);
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const columnData = sheet.getRange(1, columnIndex, sheet.getLastRow() + 1).getValues();
    for (let i = 0; i < columnData.length; i++) {
      if (columnData[i][0] === giaTriThamChieu) return i + 1;
    }
    return -1;
  }
}

