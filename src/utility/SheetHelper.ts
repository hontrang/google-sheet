/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
namespace SheetHelper {
  export const SHEET_THAM_CHIEU = 'tham chiếu';
  export const SHEET_BANG_THONG_TIN = 'bảng thông tin';
  export const SHEET_DU_LIEU = 'dữ liệu';
  export const SHEET_CHI_TIET_MA = 'chi tiết mã';
  export const SHEET_CAU_HINH = 'cấu hình';
  export const SHEET_DEBUG = 'debug';
  export const KICH_THUOC_MANG_PHU = 10;
  export const SHEET_HOSE = 'HOSE';
  export const SHEET_GIA = 'Giá';
  export const SHEET_KHOI_LUONG = 'Khối Lượng';
  export const SHEET_KHOI_NGOAI_MUA = 'KN Mua';
  export const SHEET_KHOI_NGOAI_BAN = 'KN Bán';
  export const SHEET_TY_GIA = 'Tỷ Giá USD/VND';

  export function ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(row, column, data.length, data[0].length).clearContent();
    sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  }

  export function ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    sheet.getRange(range).clearContent();
    sheet.getRange(range).setValues(data);
  }

  export function ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return;
    }
    const rowIndex = rowNumber - 1;
    const columnIndex = doiTenCotThanhChiSo(columnName) - 1;

    sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
    try {
      sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      console.error(e);
    }
  }
  /**
   *
   * @param tenMa Tìm trong duLieuCotThamChieu vị trí của TenMa, trả về vị trí cuối cùng tìm thấy
   * @param sheetName
   * @param duLieuCotThamChieu
   * @param hangBatDau
   * @returns
   */
  export function layViTriCotThamChieu(tenMa: string, sheetName: string, duLieuCotThamChieu: string[], hangBatDau: number): number {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    let vitri = -1;
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return vitri;
    }
    for (let i = 0; i < duLieuCotThamChieu.length; i++) {
      if (duLieuCotThamChieu[i] === tenMa) vitri = i + hangBatDau;
    }
    return vitri;
  }

  export function ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    sheet.getRange(cell).setValue(data);
    return true;
  }

  export function layDuLieuTrongOTheoTen(sheetName: string, cell: string): string {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return 'invalid sheet name';
    return sheet.getRange(cell).getValue();
  }

  export function layDuLieuTrongO(sheetName: string, cell: string): string {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return '';
    return sheet.getRange(cell).getValue();
  }

  export function layDuLieuTrongCot(sheetName: string, column: string): string[] {
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
    dataArray.shift();
    return dataArray;
  }

  export function laySoHangTrongSheet(sheetName: string): number {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return -1;
    return sheet.getLastRow();
  }

  export function doiTenCotThanhChiSo(columnName: string): number {
    let index = 0;
    const length = columnName.length;
    for (let i = 0; i < length; i++) {
      const charCode = columnName.toUpperCase().charCodeAt(i) - 64;
      index += charCode * Math.pow(26, length - i - 1);
    }
    return index;
  }

  export function layDuLieuTrongHang(sheetName: string, rowIndex: number): string[] {
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

  export function chen1HangVaoDauSheet(sheetName: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    sheet.insertRowsBefore(1, 1);
    return true;
  }

  // xóa cột và dời dữ liệu cột sau đó về trước
  export function xoaCot(sheetName: string, column: string, numOfCol: number): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    sheet.deleteColumns(SheetHelper.doiTenCotThanhChiSo(column), numOfCol);
    return true;
  }

  export function xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.log('Sheet không tồn tại');
      return false;
    }
    const numRows = sheet.getLastRow() - startRow + 1;
    const range = sheet.getRange(startRow, SheetHelper.doiTenCotThanhChiSo(column), numRows, numOfCol);
    range.clear();
    return true;
  }
}
