import * as luxon from 'luxon';
import { LocaleOptions } from 'luxon';

export class DateHelper {
  public static layNgay(ngay: string): string {
    const date: Date = new Date(ngay);
    return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
  }
  public static layNamHienTai(): number {
    const date: Date = new Date();
    return date.getFullYear();
  }

  public static doiDinhDangNgay(date: string, formatFrom: string, formatTo: string, opts?: LocaleOptions, zone?: string): string {
    let fromeZone = 'utc';
    let defaultZone = 'Asia/Bangkok';
    if (zone !== undefined) {
      defaultZone = zone;
    }
    return luxon.DateTime.fromFormat(date, formatFrom, { zone: fromeZone }).setZone(defaultZone).toFormat(formatTo, opts).toString();
  }
  public static doiDinhDangNgayISO(date: string, formatTo: string): string {
    return luxon.DateTime.fromISO(date).toFormat(formatTo).toString();
  }

  public static layNgayHienTaiTruSoNgay(format: string, days: number): string {
    return luxon.DateTime.now().minus({ days: days }).toFormat(format).toString();
  }
  public static layNgayHienTai(format: string): string {
    return luxon.DateTime.now().toFormat(format).toString();
  }
  public static doiTuMillisSangNgay(millis: number, format: string): string {
    return luxon.DateTime.fromMillis(millis).toFormat(format).toString();
  }

  public static layKiTaiChinhTheoQuy(): string {
    const currentDate = luxon.DateTime.now().toFormat('yyyy-MM-dd').toString();
    const month = Number(currentDate.split('-')[1]);
    const year = Number(currentDate.split('-')[0]);
    let fiscalDate = '';
    if (month < 1 || month > 12) fiscalDate = 'invalid datetime';
    else {
      if (month >= 1 && month <= 3) fiscalDate = `${year - 1}-09-30,${year - 1}-12-31`;
      if (month >= 4 && month <= 6) fiscalDate = `${year}-01-31,${year}-03-31`;
      if (month >= 7 && month <= 9) fiscalDate = `${year}-03-31,${year}-06-30`;
      if (month >= 10 && month <= 12) fiscalDate = `${year}-06-30,${year}-09-30`;
    }
    return fiscalDate;
  }
}
