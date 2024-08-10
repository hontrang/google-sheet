/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as fs from 'fs';
import path from 'path';
import { readFile, set_fs, utils, WorkBook, WorkSheet } from 'xlsx';

export class ExcelHelper {
    // eslint-disable-next-line @typescript-eslint/naming-convention
    readonly FILE_PATH = path.join(__dirname, "./dw.xlsx");
    private workBook!: WorkBook;
    private workSheet!: WorkSheet;
    constructor() {
        set_fs(fs);
        this.workBook = readFile(this.FILE_PATH);
    }
    ghiDuLieuVaoDay(data: any[][], sheetName: string, row: number, column: number): void {

        throw new Error('Method not implemented.');
    }
    ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void {
        throw new Error('Method not implemented.');
    }
    ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowNumber: number, columnName: string): void {
        throw new Error('Method not implemented.');
    }
    layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number {
        throw new Error('Method not implemented.');
    }
    ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
        throw new Error('Method not implemented.');
    }
    layDuLieuTrongOTheoTen(sheetName: string, cell: string): string {
        throw new Error('Method not implemented.');
    }
    layDuLieuTrongO(sheetName: string, cell: string) {
        const ws: WorkSheet = this.getSheetByName(sheetName);
        return ws[cell].v;
    }
    layDuLieuTrongCot(sheetName: string, column: string): string[] {
        const columnData: string[] = [];
        this.workSheet = this.getSheetByName(sheetName);
        const range = utils.decode_range(this.workSheet['!ref'] ?? "");
        for (let row = range.s.r; row <= range.e.r; row++) {
            const cellAddress = `${column}${row + 1}`;
            const cell = this.workSheet[cellAddress];
            const value = cell ? cell.v : '';
            columnData.push(value);
        }
        return columnData;
    }
    laySoHangTrongSheet(sheetName: string): number {
        throw new Error('Method not implemented.');
    }
    doiTenCotThanhChiSo(columnName: string): number {
        throw new Error('Method not implemented.');
    }
    layDuLieuTrongHang(sheetName: string, rowIndex: number): string[] {
        throw new Error('Method not implemented.');
    }
    chen1HangVaoDauSheet(sheetName: string): boolean {
        throw new Error('Method not implemented.');
    }
    xoaCot(sheetName: string, column: string, numOfCol: number): boolean {
        throw new Error('Method not implemented.');
    }
    xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean {
        throw new Error('Method not implemented.');
    }

    getSheetByName(name = "hose"): WorkSheet {
        this.workSheet = this.workBook.Sheets[name];
        return this.workSheet;
    }
}

