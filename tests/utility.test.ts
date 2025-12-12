import { test, expect } from '@playwright/test';
import { DateHelper } from '@utils/DateHelper';

test.describe('kiểm tra chức năng của DateHelper', () => {
  test('kiểm tra hàm [layNamHienTaiAsString]', async () => {
    const actual = DateHelper.layNamHienTaiAsString();
    expect(actual).toContain('2025');
  });

  test('kiểm tra hàm [doiDinhDangNgay]', async () => {
    const actual = DateHelper.doiDinhDangNgay('2025-12-18', 'yyyy-MM-dd', 'EEEE yyyy/MM/dd', { locale: 'vi-VN' });
    expect(actual).toContain('Thứ Năm 2025/12/18');
  });
});
