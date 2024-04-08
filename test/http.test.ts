const A = "111";
console.log(A);



describe("Validate url vndirect works correctly", () => {
    test("it should validate response api vndirect", () => {
        const tenMa = "HPG";
        // const fromDate = moment().format("YYYY-MM-DD");
        // const toDate = moment().format("YYYY-MM-DD");
        const fromDate = "2024-04-08";
        const toDate = "2024-04-08";
        const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
        const object = SheetHttp.sendGetRequest(URL);
        expect(object.data).not.toBeNull();
    });
});