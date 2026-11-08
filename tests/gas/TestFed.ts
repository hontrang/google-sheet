/* eslint-disable @typescript-eslint/no-unused-vars */
import { FomcHelper } from '@utils/FomcHelper';
import { FredHelper } from '@utils/FredHelper';

function TestFed() {
  QUnit.module('FOMC / FRED helper');

  QUnit.test('Test phanTichNgay', function (assert: Assert) {
    const helper = new FomcHelper();
    assert.deepEqual(
      helper.phanTichNgay(2026, 'January', '27-28'),
      { nam: 2026, ngayBatDau: '2026-01-27', ngayKetThuc: '2026-01-28', coDuBaoKinhTe: false, ghiChu: '' },
      'hai ngày'
    );
    assert.equal(helper.phanTichNgay(2026, 'Apr/May', '30-1')?.ngayKetThuc, '2026-05-01', 'qua tháng');
    assert.ok(helper.phanTichNgay(2026, 'March', '17-18*')?.coDuBaoKinhTe, 'dấu * là có SEP');
    assert.equal(helper.phanTichNgay(2026, 'March', '9 (unscheduled)')?.ghiChu, '(unscheduled)', 'ghi chú');
    assert.equal(helper.phanTichNgay(2026, 'Foo', '1-2'), undefined, 'tháng lạ');
  });
  QUnit.test('Test layLichFomc', function (assert: Assert) {
    const lich = new FomcHelper().layLichFomc();
    assert.ok(lich.length > 0, 'có cuộc họp');
    assert.ok(
      lich.every((c, i) => i === 0 || lich[i - 1].ngayKetThuc <= c.ngayKetThuc),
      'sắp xếp tăng dần'
    );
    assert.ok(new FomcHelper().layCuocHopTiepTheo() !== undefined, 'có cuộc họp tiếp theo');
  });
  QUnit.test('Test phanTichQuanSat', function (assert: Assert) {
    const helper = new FredHelper('test');
    const data = JSON.stringify({
      observations: [
        { date: '2026-11-06', value: '4.00' },
        { date: '2026-11-05', value: '.' }
      ]
    });
    assert.deepEqual(helper.phanTichQuanSat(data), [{ ngay: '2026-11-06', giaTri: 4 }], 'bỏ ngày không có số liệu');
    assert.throws(() => helper.phanTichQuanSat('{"error_code":400,"error_message":"Bad Request."}', 400), /FRED API lỗi 400/, 'lỗi API');
  });
  QUnit.test('Test layLaiSuatMucTieuFed', function (assert: Assert) {
    if (!PropertiesService.getScriptProperties().getProperty(FredHelper.tenThuocTinhApiKey)) {
      assert.ok(true, 'bỏ qua — chưa có FRED_API_KEY trong Script Properties');
      return;
    }
    const laiSuat = new FredHelper().layLaiSuatMucTieuFed();
    assert.ok(laiSuat !== undefined && laiSuat.canTren > laiSuat.canDuoi, 'cận trên > cận dưới');
  });
}
