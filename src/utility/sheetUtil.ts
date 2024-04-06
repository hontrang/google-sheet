import { SheetLog } from "./logUtil";
declare const SpreadsheetApp: any;

interface SheetUtilityType {
    SHEET_THAM_CHIEU: string;
    SHEET_BANG_THONG_TIN: string;
    SHEET_DU_LIEU: string;
    SHEET_CHI_TIET_MA: string;
    SHEET_CAU_HINH: string;
    SHEET_DEBUG: string;
    KICH_THUOC_MANG_PHU: number;
    ghiDuLieuVaoDay: (data: any[][], sheetName: string, row: number, column: number) => void;
    ghiDuLieuVaoDayTheoTen: (data: any[][], sheetName: string, rowNumber: number, columnName: string) => void;
    ghiDuLieuVaoO: (data: any, sheetName: string, cell: string) => void;
    layDuLieuTrongO: (sheetName: string, cell: string) => any;
    layDuLieuTrongCot: (sheetName: string, column: string) => any[];
    columnToIndex: (columnName: string) => number;
    chen1HangVaoDauSheet: (sheetName: string) => number;
}

const SheetUtil: SheetUtilityType = {
    SHEET_THAM_CHIEU: "tham chiếu",
    SHEET_BANG_THONG_TIN: "bảng thông tin",
    SHEET_DU_LIEU: "dữ liệu",
    SHEET_CHI_TIET_MA: "chi tiết mã",
    SHEET_CAU_HINH: "cấu hình",
    SHEET_DEBUG: "debug",
    KICH_THUOC_MANG_PHU: 10,
    ghiDuLieuVaoDay(data, sheetName, row, column) {
        let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return -1;
        }
        sheet.getRange(row, column, data.length, data[0].length).clearContent();
        sheet.getRange(row, column, data.length, data[0].length).setValues(data);
    },
    ghiDuLieuVaoDayTheoTen(data, sheetName, rowNumber, columnName) {
        let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return -1;
        }
        let rowIndex = rowNumber - 1;
        let columnIndex = this.columnToIndex(columnName) - 1;

        sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
        try {
            sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
        } catch (e) {
            SheetLog.logDebug(data);
            console.error(e);
        }
    },
    ghiDuLieuVaoO(data, sheetName, cell) {
        SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).setValue(data);
    },
    layDuLieuTrongO(sheetName, cell) {
        return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(cell).getValue();
    },
    layDuLieuTrongCot(sheetName, column) {
        const columnData = SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(`${column}:${column}`).getValues();

        const dataArray: any[] = [];
        for (const element of columnData) {
            const value = element[0];
            if (value !== "") {
                dataArray.push(value);
            }
        }
        dataArray.shift(); // Xóa phần tử đầu tiên ("title") trong mảng dataArray
        return dataArray;
    },
    columnToIndex(columnName) {
        let index = 0;
        let length = columnName.length;
        for (let i = 0; i < length; i++) {
            let charCode = columnName.toUpperCase().charCodeAt(i) - 64;
            index += (charCode * Math.pow(26, length - i - 1));
        }
        return index;
    },
    chen1HangVaoDauSheet(sheetName) {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return -1;
        }
        sheet.insertRowsBefore(1, 1);
        return 1;
    }
};
export { SheetUtil }