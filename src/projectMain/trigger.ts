/* eslint-disable @typescript-eslint/no-unused-vars */
/**
 * All triggers are executed manually or via clasp
 */

/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function mainCreateTriggers(): void {
  // Tạo trigger chạy theo thời gian
  ScriptApp.newTrigger('getDataHose').timeBased().everyHours(1).create();

  ScriptApp.newTrigger('layTinTucSheetBangThongTin').timeBased().everyDays(1).atHour(4).nearMinute(0).create();

  ScriptApp.newTrigger('batSukienSuaThongTinO').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();

  ScriptApp.newTrigger('getDataHose').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  ScriptApp.newTrigger('vietstock').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('derivative').timeBased().everyMinutes(5).create();

  Logger.log('Các trigger đã được tạo!');
}

/**
 * Deletes all triggers in the project.
 * @see https://developers.google.com/apps-script/guides/triggers/installable
 */
function mainDeleteTrigger(): void {
  // Lấy tất cả triggers trong dự án
  const allTriggers: GoogleAppsScript.Script.Trigger[] = ScriptApp.getProjectTriggers();

  // Sử dụng vòng lặp for-of để duyệt qua mỗi trigger
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}
