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
  // Lấy tất cả triggers trong dự án
  const allTriggers = ScriptApp.getProjectTriggers();
  
  // Sử dụng vòng lặp for-of để duyệt qua mỗi trigger
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}