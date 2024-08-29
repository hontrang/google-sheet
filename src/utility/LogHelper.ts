/* eslint-disable @typescript-eslint/no-extraneous-class */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { SheetHelper } from "@src/utility/SheetHelper";
import { DateHelper } from "@src/utility/DateHelper";

export class LogHelper {
  public static logDebug(content: string): boolean {
    const data: [string, string][] = [[DateHelper.layNgayHienTai('YYYY/MM/DD HH:mm:ss'), content.toString()]];
    Logger.log(content);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.sheetName.sheetDebug);
    if (!sheet) return false;
    else {
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
      return true;
    }
  }

  public static logTime(sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    else {
      sheet.getRange(cell).setValue(DateHelper.layNgayHienTai('YYYY/MM/DD HH:mm:ss'));
      return true;
    }
  }

  public static logStart(sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    else {
      sheet.getRange(cell).setValue('...');
      return true;
    }
  }

  public static async sleepSync(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
