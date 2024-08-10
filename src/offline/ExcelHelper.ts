/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import path from 'path';
import * as ExcelJS from 'exceljs';
import { Workbook, Worksheet } from 'exceljs';

export class ExcelHelper {
    // eslint-disable-next-line @typescript-eslint/naming-convention
    readonly FILE_PATH = path.join(__dirname, "./dw.xlsx");
    private workBook!: Workbook;
    private workSheet!: Worksheet;

    private async readExcelFileSync(filePath: string): Promise<Workbook> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath).then(() => {
            console.log('Workbook đã được nạp xong.');
        }).catch(error => {
            console.error('Lỗi khi nạp workbook:', error);
        });
        return workbook;
    }

    public async initializeWorkbook() {
        this.workBook = await this.readExcelFileSync(this.FILE_PATH);
    }

    getSheetByName(name = "hose"): Worksheet {
        let sheetID;
        this.workBook.eachSheet(function (sheet, id) {
            if (name === sheet.name) sheetID = id;
        });
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        return this.workBook.getWorksheet(sheetID)!;
    }

    layDuLieuTrongO(sheetName: string, cell: string): string {
        this.workSheet = this.getSheetByName(sheetName);
        return this.workSheet.getCell(cell).text;
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
    layDuLieuTrongCot(sheetName: string, column: string): string[] {
        const columnData: string[] = [];
        this.workSheet = this.getSheetByName(sheetName);
        this.workSheet.getColumn(column).eachCell(function (cell) {
            columnData.push(cell.text);
        });
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


}

