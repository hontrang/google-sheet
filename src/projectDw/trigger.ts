/* eslint-disable @typescript-eslint/no-unused-vars */
/**
 * All triggers are executed manually or via clasp
 */

/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function dwCreateTriggers() {
  // Tạo trigger chạy theo thời gian
  ScriptApp.newTrigger('layChiSoVnIndex').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('layThongTinCoBan').timeBased().everyDays(1).atHour(18).nearMinute(0).create();
  ScriptApp.newTrigger('layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay').timeBased().everyDays(1).atHour(15).nearMinute(50).create();
  ScriptApp.newTrigger('LAY_SU_KIEN').timeBased().everyDays(1).atHour(1).create();
  ScriptApp.newTrigger('layThongTinPhaiSinh').timeBased().everyDays(1).atHour(1).create();
  ScriptApp.newTrigger('layBaoCaoDC').timeBased().everyMinutes(5).create();
  Logger.log('Các trigger đã được tạo!');
}

/**
 * Deletes a trigger.
 * @param {string} triggerId The Trigger ID.
 * @see https://developers.google.com/apps-script/guides/triggers/installable
 */
function dwDeleteTrigger() {
  // Loop over all triggers.
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const element of allTriggers) {
    ScriptApp.deleteTrigger(element);
  }
}
