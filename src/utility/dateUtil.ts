namespace DateUtil {
    export function getDate(num: string) {
        let date: Date = new Date(num);
        return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
    }
}