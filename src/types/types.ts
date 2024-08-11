/* eslint-disable @typescript-eslint/naming-convention */
/* eslint-disable @typescript-eslint/no-explicit-any */
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
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

export interface Http {
        sendRequest(url: string, option?: URLFetchRequestOptions): any;
        sendPostRequest(url: string, options?: URLFetchRequestOptions): any;
        sendGetRequest(url: string): any;
        sendGraphQLRequest(url: string, query: string, variables?: any): any;
        getToken(): string;
}

export interface ResponseVndirect {
        code?: string,
        type?: string,
        tradingDate?: string,
        floor?: string,
        buyVal?: number,
        sellVal?: number,
        netVal?: number,
        buyVol?: number,
        sellVol?: number,
        netVol?: number,
        totalRoom?: number,
        currentRoom?: number,
        nmVolume?: number
}

export interface ResponseSsi {
        Symbol?: string,
        ClosePrice?: number
}