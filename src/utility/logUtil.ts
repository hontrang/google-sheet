import moment from 'moment';

declare const SpreadsheetApp: any;
declare const SheetUtility: any;

const SheetLog = {
  logDebug: function (content: any): void {
    const data: [string, string][] = [
      [moment().format("YYYY/MM/DD HH:mm:ss"), content.toString()]
    ];
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_DEBUG);
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    Logger.log(content);
  },
  logTime: function (sheetName: string, cell: string): void {
    SpreadsheetApp.getActive()
      .getSheetByName(sheetName)
      .getRange(cell)
      .setValue(moment().format("YYYY/MM/DD HH:mm:ss"));
  },

  logStart: function (sheetName: string, cell: string): void {
    SpreadsheetApp.getActive()
      .getSheetByName(sheetName)
      .getRange(cell)
      .setValue("...");
  }
};
export { SheetLog }