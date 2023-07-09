var SheetDate = {
    getDate: function (number) {
        let date = new Date(number);
        return (date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate());
    }
}