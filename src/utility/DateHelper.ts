// eslint-disable-next-line @typescript-eslint/no-unused-vars

export class DateHelper {
  public static layNgay(ngay: string): string {
    const date: Date = new Date(ngay);
    return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
  }

  public static doiDinhDangNgay(date: string, formatFrom: string, formatTo: string): string {
    // return moment(date, formatFrom, true).format(formatTo);
    return date;
  }

  public static layNgayHienTai(format: string): string {
    //return moment().format(format);
    return "Done";
  }
}
