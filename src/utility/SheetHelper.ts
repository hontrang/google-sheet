/* eslint-disable @typescript-eslint/no-explicit-any */

import { ISheetSpread } from '@src/types/generic';

// eslint-disable-next-line @typescript-eslint/no-unused-vars, @typescript-eslint/no-extraneous-class
export class SheetHelper implements ISheetSpread {
  public static readonly sheetName = {
    sheetThamChieu: 'tham chiếu',
    sheetBangThongTin: 'bảng thông tin',
    sheetDuLieu: 'dữ liệu',
    sheetChiTietMa: 'chi tiết mã',
    sheetCauHinh: 'cấu hình',
    sheetDebug: 'debug',
    sheetHose: 'hose',
    sheetGia: 'Giá',
    sheetKhoiLuong: 'Khối Lượng',
    sheetKhoiNgoaiMua: 'KN Mua',
    sheetKhoiNgoaiBan: 'KN Bán',
    sheetTyGia: 'Tỷ Giá USD/VND',
    sheetDC: 'DC'
  };

  public static readonly kichThuocMangPhu = 10;
  layDuLieuTrongO(sheetName: string, cell: string): string {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return '';
    return sheet.getRange(cell).getValue();
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

  layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number {
    return duLieuCotThamChieu.indexOf(tenMa) + hangBatDau;
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
    return dataArray;
  }

  laySoHangTrongSheet(sheetName: string): number {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return -1;
    return sheet.getLastRow();
  }

  taoSheetMoi(sheetName: string) {
    SpreadsheetApp.getActive().insertSheet(sheetName);
  }

  kiemTraSheetTonTai(sheetName: string) {
    const sheetApp = SpreadsheetApp.getActive();
    const sheet = sheetApp.getSheetByName(sheetName);
    return sheet !== null;
  }

  xoaSheet(sheetName: string) {
    const sheetApp = SpreadsheetApp.getActive();
    const sheet = sheetApp.getSheetByName(sheetName);
    if (sheet) {
      sheetApp.deleteSheet(sheet);
    }
  }

  ghiDuLieuVaoO(data: unknown, sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    sheet.getRange(cell).setValue(data);
    return true;
  }

  ghiDuLieuVaoDay(data: unknown[][], sheetName: string, row: number, column: number): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(row, column, data.length, data[0].length).clearContent();
    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  }

  ghiDuLieuVaoDayTheoVung(data: unknown[][], sheetName: string, range: string): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(range).clearContent();
    sheet.getRange(range).setValues(data);
  }

  ghiDuLieuVaoDayTheoTen(data: unknown[][], sheetName: string, rowNumber: number, columnName: string): void {
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

  doiTenCotThanhChiSo(columnName: string): number {
    let index = 0;
    const length = columnName.length;
    for (let i = 0; i < length; i++) {
      const charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += charCode * Math.pow(26, length - i - 1);
    }
    return index;
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

  xoaHang(sheetName: string, rowIndex: number) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return;
    sheet.deleteRow(rowIndex);
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

  xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number, endRow?: number): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    let numRows;
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    if (endRow !== undefined) {
      numRows = endRow;
    } else {
      numRows = sheet.getLastRow() - startRow + 1;
    }
    const range = sheet.getRange(startRow, this.doiTenCotThanhChiSo(column), numRows, numOfCol);
    range.clearContent();
    return true;
  }
}
