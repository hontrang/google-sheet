import * as luxon from "luxon";

export class DateHelper {
  public static layNgay(ngay: string): string {
    const date: Date = new Date(ngay);
    return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
  }

  public static doiDinhDangNgay(date: string, formatFrom: string, formatTo: string): string {
    return luxon.DateTime.fromFormat(date, formatFrom).toFormat(formatTo).toString();
  }

  public static layNgayHienTai(format: string): string {
    return luxon.DateTime.now().toFormat(format).toString();
  }
}
