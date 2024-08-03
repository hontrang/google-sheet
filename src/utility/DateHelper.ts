// eslint-disable-next-line @typescript-eslint/no-unused-vars
namespace DateHelper {
  export function getDate(num: string): string {
    const date: Date = new Date(num);
    return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
  }

  export function changeFormatDate(date: string, formatFrom: string, formatTo: string): string {
    return moment(date, formatFrom, true).format(formatTo);
  }
  export function layNgayHienTai(format: string) {
    return moment().format(format);
  }
}
