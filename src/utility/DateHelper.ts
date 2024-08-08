// eslint-disable-next-line @typescript-eslint/no-unused-vars
import moment from "moment";
namespace DateHelper {
  export function layNgay(ngay: string): string {
    const date: Date = new Date(ngay);
    return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
  }

  export function doiDinhDangNgay(date: string, formatFrom: string, formatTo: string): string {
    return moment(date, formatFrom, true).format(formatTo);
  }
  export function layNgayHienTai(format: string) {
    return moment().format(format);
  }
}
