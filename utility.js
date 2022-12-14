function logTime(sheet, cell) {
    SpreadsheetApp.getActive().getSheetByName(sheet).getRange(cell).setValue(new Date());
  }
  