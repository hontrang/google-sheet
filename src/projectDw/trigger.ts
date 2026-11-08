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
  ScriptApp.newTrigger('layTinTucHSX').timeBased().everyDays(2).atHour(1).create();
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

/**
 * Xoá các trigger gọi tới một hàm, không đụng trigger khác: `clasp run dwXoaTriggerTheoHam --params '["LAY_SU_KIEN"]'`.
 * Trả về số trigger đã xoá.
 */
function dwXoaTriggerTheoHam(tenHam: string): number {
  let soLuong = 0;
  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction() === tenHam) {
      ScriptApp.deleteTrigger(trigger);
      soLuong++;
    }
  }
  return soLuong;
}

/**
 * Liệt kê các trigger đang đăng ký trên GAS (chạy qua `clasp run dwListTriggers` để kiểm tra).
 */
function dwListTriggers(): string[] {
  return ScriptApp.getProjectTriggers().map((trigger) => `${trigger.getHandlerFunction()} | ${trigger.getEventType().toString()} | ${trigger.getUniqueId()}`);
}
