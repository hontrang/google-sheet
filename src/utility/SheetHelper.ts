/* eslint-disable @typescript-eslint/no-explicit-any */

import { SheetSpread } from "../types/types";

// eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-extraneous-class
export class SheetHelper implements SheetSpread {
  static readonly SheetName = {
    SHEET_THAM_CHIEU: 'tham chiếu',
    SHEET_BANG_THONG_TIN: 'bảng thông tin',
    SHEET_DU_LIEU: 'dữ liệu',
    SHEET_CHI_TIET_MA: 'chi tiết mã',
    SHEET_CAU_HINH: 'cấu hình',
    SHEET_DEBUG: 'debug',
    SHEET_HOSE: 'HOSE',
    SHEET_GIA: 'Giá',
    SHEET_KHOI_LUONG: 'Khối Lượng',
    SHEET_KHOI_NGOAI_MUA: 'KN Mua',
    SHEET_KHOI_NGOAI_BAN: 'KN Bán',
    SHEET_TY_GIA: 'Tỷ Giá USD/VND',
  };

  public static readonly KICH_THUOC_MANG_PHU = 10;

  ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(row, column, data.length, data[0].length).clearContent();
    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  }

  ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(range).clearContent();
    sheet.getRange(range).setValues(data);
  }

  ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    const rowIndex = rowNumber - 1;
    const columnIndex = this.doiTenCotThanhChiSo(columnName) - 1;

    sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
    try {
      sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      console.error(e);
    }
  }

  layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number {
    let vitri = -1;
    for (let i = 0; i < duLieuCotThamChieu.length; i++) {
      if (duLieuCotThamChieu[i] === tenMa) vitri = i + hangBatDau;
    }
    return vitri;
  }

  ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    sheet.getRange(cell).setValue(data);
    return true;
  }

  layDuLieuTrongO(sheetName: string, cell: string): string {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return '';
    return sheet.getRange(cell).getValue();
  }

  layDuLieuTrongCot(sheetName: string, column: string): string[] {
    const dataArray: string[] = [];
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return dataArray;
    const columnData = sheet.getRange(`${column}:${column}`).getValues();

    for (const element of columnData) {
      const value = element[0];
      if (value !== '') {
        dataArray.push(value);
      }
    }
    // bỏ hàng đầu là tên cột
    dataArray.shift();
    return dataArray;
  }

  laySoHangTrongSheet(sheetName: string): number {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return -1;
    return sheet.getLastRow();
  }

  doiTenCotThanhChiSo(columnName: string): number {
    let index = 0;
    const length = columnName.length;
    for (let i = 0; i < length; i++) {
      const charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += charCode * Math.pow(26, length - i - 1);
    }
    return index;
  }

  layDuLieuTrongHang(sheetName: string, rowIndex: number): string[] {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return [];
    // Lấy số lượng cột trong Sheet
    const numColumns = sheet.getLastColumn();

    // Lấy dữ liệu từ hàng
    const range = sheet.getRange(rowIndex, 1, 1, numColumns);
    const rowData = range.getValues();

    // rowData là một mảng 2 chiều, chúng ta cần phải lấy phần tử đầu tiên để có mảng 1 chiều
    return rowData[0];
  }

  chen1HangVaoDauSheet(sheetName: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    sheet.insertRowsBefore(1, 1);
    return true;
  }

  xoaCot(sheetName: string, column: string, numOfCol: number): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    sheet.deleteColumns(this.doiTenCotThanhChiSo(column), numOfCol);
    return true;
  }

  xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    const numRows = sheet.getLastRow() - startRow + 1;
    const range = sheet.getRange(startRow, this.doiTenCotThanhChiSo(column), numRows, numOfCol);
    range.clear();
    return true;
  }
}
