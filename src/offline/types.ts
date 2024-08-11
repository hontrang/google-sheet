/* eslint-disable @typescript-eslint/no-explicit-any */
export interface SheetSpread {
        layDuLieuTrongO(sheetName: string, cell: string): string;
        layDuLieuTrongCot(sheetName: string, column: string): string[];
        laySoHangTrongSheet(sheetName: string): number;
        layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number;
        layDuLieuTrongHang(sheetName: string, rowIndex: number): string[];
        ghiDuLieuVaoDay(data: any[][], sheetName: string, rowIndex: number, columnIndex: number): void;
        ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void;
        ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void;
        ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean;
        doiTenCotThanhChiSo(columnName: string): number;
        chen1HangVaoDauSheet(sheetName: string): boolean;
        xoaCot(sheetName: string, column: string, numOfCol: number): boolean;
        xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean;
}