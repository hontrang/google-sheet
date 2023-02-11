var SheetLog = {
    logDebug: function (content) {
      let data = [
        [moment().format("YYYY/MM/DD HH:mm:ss"), content.toString()]
      ];
      SpreadsheetApp.getActive().getSheetByName(SHEET_DEBUG).getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
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