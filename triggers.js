/**
 * All triggers are executed manually or via clasp
 */

/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function createTriggers() {
  // Tạo trigger chạy theo thời gian
  ScriptApp.newTrigger("getDataHose")
    .timeBased()
    .everyHours(1)
    .create();
    
  ScriptApp.newTrigger("layTinTucSheetBangThongTin")
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .nearMinute(0)
    .create();

  ScriptApp.newTrigger("batSukienSuaThongTinO")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  
  ScriptApp.newTrigger("getDataHose")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  Logger.log("Các trigger đã được tạo!");
}

/**
 * Deletes a trigger.
 * @param {string} triggerId The Trigger ID.
 * @see https://developers.google.com/apps-script/guides/triggers/installable
 */
function deleteTrigger() {
  // Loop over all triggers.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let index = 0; index < allTriggers.length; index++) {
    ScriptApp.deleteTrigger(allTriggers[index]);
  }
}