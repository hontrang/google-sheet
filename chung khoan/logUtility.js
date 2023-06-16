var SheetLog = {
    logDebug: function (content) {
      const data = [
        [moment().format("YYYY/MM/DD HH:mm:ss"), content.toString()]
      ];
      var sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_DEBUG);
      sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
      Logger.log(content);
    },
    logTime: function (sheetName, cell) {
      SpreadsheetApp.getActive()
        .getSheetByName(sheetName)
        .getRange(cell)
        .setValue(moment().format("YYYY/MM/DD HH:mm:ss"));
    },
  
    logStart: function (sheetName, cell) {
      SpreadsheetApp.getActive()
        .getSheetByName(sheetName)
        .getRange(cell)
        .setValue("...");
    }
  };