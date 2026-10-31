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
  taoTriggerGhiLaiLoCuoiNam();

  Logger.log('Các trigger đã được tạo!');
}

/**
 * Tạo (lại) trigger ghi lãi/lỗ vào sheet `vốn` lúc ~23h30 ngày 31/12 hằng năm.
 * Chạy riêng được mà không đụng tới các trigger khác.
 */
function taoTriggerGhiLaiLoCuoiNam(): void {
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === 'ghiLaiLoCuoiNam') ScriptApp.deleteTrigger(trigger);
  }
  ScriptApp.newTrigger('ghiLaiLoCuoiNam').timeBased().onMonthDay(31).atHour(23).nearMinute(30).create();
}

/**
 * Liệt kê các trigger đang đăng ký trên GAS (chạy qua `clasp run mainListTriggers` để kiểm tra).
 */
function mainListTriggers(): string[] {
  return ScriptApp.getProjectTriggers().map((trigger) => `${trigger.getHandlerFunction()} | ${trigger.getEventType().toString()} | ${trigger.getUniqueId()}`);
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
