interface Sheet{
        ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void;
    
       ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void;
    
       ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void;
    
       layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number;
    
       ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean;
    
       layDuLieuTrongOTheoTen(sheetName: string, cell: string): string;
    
       layDuLieuTrongO(sheetName: string, cell: string): string;
    
       layDuLieuTrongCot(sheetName: string, column: string): string[];
    
       laySoHangTrongSheet(sheetName: string): number;
    
       doiTenCotThanhChiSo(columnName: string): number;
    
       layDuLieuTrongHang(sheetName: string, rowIndex: number): string[];
    
       chen1HangVaoDauSheet(sheetName: string): boolean;
    
       xoaCot(sheetName: string, column: string, numOfCol: number): boolean;
    
       xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean;
}