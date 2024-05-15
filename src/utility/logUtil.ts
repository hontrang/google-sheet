class SheetLog {
    static logDebug(content: string): boolean {
        const data: [string, string][] = [
            [moment().format("YYYY/MM/DD HH:mm:ss"), content]
        ];
        Logger.log(content);
        const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtil.SHEET_DEBUG);
        if (!sheet) return false;
        else {
            sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
            return true;
        }
    }

    static logTime(sheetName: string, cell: string): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return false;
        else {
            sheet.getRange(cell).setValue(moment().format("YYYY/MM/DD HH:mm:ss"));
            return true;
        }
    }

    static logStart(sheetName: string, cell: string): boolean {
        const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (!sheet) return false;
        else {
            sheet.getRange(cell).setValue("...");
            return true;
        }
    }
}