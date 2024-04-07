namespace SheetUtil {
    export const SHEET_THAM_CHIEU = "tham chiếu";
    export const SHEET_BANG_THONG_TIN = "bảng thông tin";
    export const SHEET_DU_LIEU = "dữ liệu";
    export const SHEET_CHI_TIET_MA = "chi tiết mã";
    export const SHEET_CAU_HINH = "cấu hình";
    export const SHEET_DEBUG = "debug";
    export const KICH_THUOC_MANG_PHU = 10;

    export function ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return;
        }
        sheet.getRange(row, column, data.length, data[0].length).clearContent();
        sheet.getRange(row, column, data.length, data[0].length).setValues(data);
    }

    export function ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return;
        }
        const rowIndex = rowNumber - 1;
        const columnIndex = columnToIndex(columnName) - 1;

        sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
        try {
            sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
        } catch (e) {
            console.error(e);
        }
    }

    export function ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return false;
        sheet.getRange(cell).setValue(data);
        return true;
    }

    export function layDuLieuTrongO(sheetName: string, cell: string): string {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return "";
        return sheet.getRange(cell).getValue();
    }

    export function layDuLieuTrongCot(sheetName: string, column: string): string[] {
        const dataArray: string[] = [];
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return dataArray;
        const columnData = sheet.getRange(`${column}:${column}`).getValues();

        for (const element of columnData) {
            const value = element[0];
            if (value !== "") {
                dataArray.push(value);
            }
        }
        dataArray.shift(); // Xóa phần tử đầu tiên ("title") trong mảng dataArray
        return dataArray;
    }

    export function columnToIndex(columnName: string): number {
        let index = 0;
        let length = columnName.length;
        for (let i = 0; i < length; i++) {
            let charCode = columnName.toUpperCase().charCodeAt(i) - 64;
            index += (charCode * Math.pow(26, length - i - 1));
        }
        return index;
    }

    export function chen1HangVaoDauSheet(sheetName: string): number {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return -1;
        }
        sheet.insertRowsBefore(1, 1);
        return 1;
    }
}