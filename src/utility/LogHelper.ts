// eslint-disable-next-line @typescript-eslint/no-unused-vars
namespace LogHelper {
  export function logDebug(content: string): boolean {
    const data: [string, string][] = [[moment().format('YYYY/MM/DD HH:mm:ss'), content.toString()]];
    Logger.log(content);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SHEET_DEBUG);
    if (!sheet) return false;
    else {
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
      return true;
    }
  }

  export function logTime(sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    else {
      sheet.getRange(cell).setValue(moment().format('YYYY/MM/DD HH:mm:ss'));
      return true;
    }
  }

  export function logStart(sheetName: string, cell: string): boolean {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return false;
    else {
      sheet.getRange(cell).setValue('...');
      return true;
    }
  }
}
