export const DateUtil: { getDate: (number: number) => string } = {
    getDate: function (number: number): string {
        let date: Date = new Date(number);
        return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`;
    }
}