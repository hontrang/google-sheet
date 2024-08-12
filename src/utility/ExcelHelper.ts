/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import path from 'path';
import * as ExcelJS from 'exceljs';
import { Cell, Workbook, Worksheet } from 'exceljs';
import { SheetSpread } from '../types/types';
import { Configuration } from 'src/configuration/Configuration';

export class ExcelHelper implements SheetSpread {
    readonly filePath = path.resolve(process.cwd(), Configuration.xlsxInput);
    readonly outPath = path.resolve(process.cwd(), Configuration.xlsxOutput);
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

    public async truyCapWorkBook() {
        this.workBook = await this.readExcelFileSync(this.filePath);
    }
    public async luuWorkBook() {
        await this.workBook.xlsx.writeFile(this.outPath).then(() => {
            console.log('Workbook đã được lưu xong.');
        }).catch(error => {
            console.error('Lỗi khi lưu workbook:', error);
        });;
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
    layViTriCotThamChieu(tenMa: string, duLieuCotThamChieu: string[], hangBatDau: number): number {
        let vitri = -1;
        for (let i = 0; i < duLieuCotThamChieu.length; i++) {
            if (duLieuCotThamChieu[i] === tenMa) vitri = i + hangBatDau;
        }
        return vitri;
    }
    layDuLieuTrongCot(sheetName: string, column: string): string[] {
        const columnData: string[] = [];
        this.workSheet = this.getSheetByName(sheetName);
        this.workSheet.getColumn(column).eachCell(function (cell) {
            const text = cell.text;
            if (text !== '') {
                columnData.push(cell.text);
            }
        });
        return columnData;
    }
    async laySoHangTrongSheet(sheetName: string): Promise<number> {
        this.workSheet = this.getSheetByName(sheetName);
        return this.workSheet.rowCount ?? 0;
    }
    layDuLieuTrongHang(sheetName: string, rowIndex: number): string[] {
        this.workSheet = this.getSheetByName(sheetName);
        const data: string[] = [];
        this.workSheet.getRow(rowIndex).eachCell(function (cell: Cell) {
            data.push(cell.text);
        });
        return data;
    }

    async ghiDuLieuVaoDay(data: any[], sheetName: string, rowIndex: number, columnIndex: number): Promise<void> {
        this.workSheet = this.getSheetByName(sheetName);
        const row = this.workSheet.getRow(rowIndex);
        const adjustedData = new Array(columnIndex - 1).fill('').concat(data);
        row.values = adjustedData;
        console.log(`Da ghi du lieu vao hang ${rowIndex} cot: ${columnIndex} data ${data}`);
        row.commit();
    }
    ghiDuLieuVaoDayTheoVung(data: any[][], sheetName: string, range: string): void {
        this.workSheet = this.getSheetByName(sheetName);
        const [startCell, endCell] = range.split(':');
        const startCellCoord: Cell = this.workSheet.getCell(startCell);
        const endCellCoord: Cell = this.workSheet.getCell(endCell);

        const startRow = startCellCoord.fullAddress.row;
        const startCol = startCellCoord.fullAddress.col;
        const endRow = endCellCoord.fullAddress.row;
        const endCol = endCellCoord.fullAddress.col;

        if (data.length !== endRow - startRow + 1 || data[0].length !== endCol - startCol + 1) {
            throw new Error('Kích thước của dữ liệu không khớp với kích thước của range.');
        }

        for (let i = 0; i < data.length; i++) {
            for (let j = 0; j < data[i].length; j++) {
                const cell = this.workSheet.getCell(startRow + i, startCol + j);
                cell.value = data[i][j];
            }
        }
    }
    ghiDuLieuVaoDayTheoTen(data: any[][], sheetName: string, rowIndex: number, columnName: string): void {
        const columnIndex = this.doiTenCotThanhChiSo(columnName);
        this.ghiDuLieuVaoDay(data, sheetName, rowIndex + 1, columnIndex);
    }

    ghiDuLieuVaoO(data: any, sheetName: string, cell: string): boolean {
        this.workSheet = this.getSheetByName(sheetName);
        this.workSheet.getCell(cell).value = data;
        return true;
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
        this.workSheet = this.getSheetByName(sheetName);
        this.workSheet.spliceRows(1, 0, []);
        return true;
    }
    xoaCot(sheetName: string, column: string, numOfCol: number): boolean {
        this.workSheet = this.getSheetByName(sheetName);
        this.workSheet.spliceColumns(this.doiTenCotThanhChiSo(column), numOfCol);
        return true;
    }
    xoaDuLieuTrongCot(sheetName: string, column: string, numOfCol: number, startRow: number): boolean {
        this.workSheet = this.getSheetByName(sheetName);
        const lastRow = this.workSheet.lastRow?.number ?? 0;

        for (let i = startRow; i <= lastRow; i++) {
            for (let j = this.doiTenCotThanhChiSo(column); j <= this.doiTenCotThanhChiSo(column) + numOfCol; j++) {
                const cell = this.workSheet.getCell(i, j);
                cell.value = null;
            }
        }
        return true;
    }
}

