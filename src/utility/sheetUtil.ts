class SheetUtil {
    static SHEET_THAM_CHIEU = "tham chiếu";
    static SHEET_BANG_THONG_TIN = "bảng thông tin";
    static SHEET_DU_LIEU = "dữ liệu";
    static SHEET_CHI_TIET_MA = "chi tiết mã";
    static SHEET_CAU_HINH = "cấu hình";
    static SHEET_DEBUG = "debug";
    static KICH_THUOC_MANG_PHU = 10;
    static SHEET_HOSE = "HOSE";
    static SHEET_GIA = "Giá";
    static SHEET_KHOI_LUONG = "Khối Lượng";
    static SHEET_KHOI_NGOAI_MUA = "KN Mua";
    static SHEET_KHOI_NGOAI_BAN = "KN Bán";
    static SHEET_TY_GIA = "Tỷ Giá USD/VND";

    static ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return;
        }
        sheet.getRange(row, column, data.length, data[0].length).clearContent();
        sheet.getRange(row, column, data.length, data[0].length).setValues(data);
    }

    static ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return;
        }
        const rowIndex = rowNumber - 1;
        const columnIndex = SheetUtil.doiTenCotThanhChiSo(columnName) - 1;

        sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).clearContent();
        try {
            sheet.getRange(rowIndex + 1, columnIndex + 1, data.length, data[0].length).setValues(data);
        } catch (e) {
            console.error(e);
        }
    }

    static ghiDuLieuVaoDayTheoTenThamChieu(data: any, sheetName: string, cotGhiDuLieu: string, cotThamChieu: string, hangBatDau: number, tenMa: string): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        let vitri = -1;
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return;
        }
        const duLieuCotThamChieu: string[] = SheetUtil.layDuLieuTrongCot(sheetName, cotThamChieu);
        for (let i = 0; i < duLieuCotThamChieu.length; i++) {
            if (duLieuCotThamChieu[i] === tenMa) vitri = i + hangBatDau;
        }
        SheetUtil.ghiDuLieuVaoO(data, sheetName, cotGhiDuLieu + vitri);
    }

    static ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return false;
        sheet.getRange(cell).setValue(data);
        return true;
    }

    static layDuLieuTrongOTheoTen(sheetName: string, cell: string): string {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return "invalid sheet name";
        return sheet.getRange(cell).getValue();
    }

    static layDuLieuTrongO(sheetName: string, cell: string): string {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return "";
        return sheet.getRange(cell).getValue();
    }

    static layDuLieuTrongCot(sheetName: string, column: string): string[] {
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
        dataArray.shift();  // Xóa phần tử đầu tiên nếu cần bỏ qua tiêu đề
        return dataArray;
    }

    static laySoHangTrongSheet(sheetName: string): number {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return -1;
        return sheet.getLastRow();
    }

    static doiTenCotThanhChiSo(columnName: string): number {
        let index = 0;
        const length = columnName.length;
        for (let i = 0; i < length; i++) {
            const charCode = columnName.toUpperCase().charCodeAt(i) - 64;
            index += charCode * Math.pow(26, length - i - 1);
        }
        return index;
    }

    static layDuLieuTrongHang(sheetName: string, rowIndex: number): string[] {
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

    static chen1HangVaoDauSheet(sheetName: string): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return false;
        }
        sheet.insertRowsBefore(1, 1);
        return true;
    }

    // Xóa cột và dời dữ liệu cột sau đó về trước
    static xoaCot(sheetName: string, column: string, numOfCol: number): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return false;
        }
        sheet.deleteColumns(SheetUtil.doiTenCotThanhChiSo(column), numOfCol);
        return true;
    }

    static xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) {
            console.log("Sheet không tồn tại");
            return false;
        }
        const numRows = sheet.getLastRow() - startRow + 1;
        const range = sheet.getRange(startRow, SheetUtil.doiTenCotThanhChiSo(column), numRows, numOfCol);
        range.clear();
        return true;
    }
}
