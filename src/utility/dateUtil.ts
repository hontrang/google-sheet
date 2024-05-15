class DateUtil {
    static getDate(num: string): string {
        const date: Date = new Date(num);
        return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
    }
    static changeFormatDate(date: string, formatFrom: string, formatTo: string): string {
        return moment(date, formatFrom, true).format(formatTo);
    }
}
