


const axios = require('axios');
describe("kiểm tra url vndirect chạy chính xác", () => {
    test("kiểm tra phản hồi từ api vndirect", async () => {
        const tenMa = "HPG";
        const fromDate = "2024-04-08";
        const toDate = "2024-04-08";
        const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
        const response = await axios.get(URL);
        const data = response.data;
        expect(data).not.toBeNull();
    });
});

describe("kiểm tra url vps chạy chính xác", () => {
    test("kiểm tra phản hồi từ api vps", async () => {
        const ma1 = "HPG";
        const ma2 = "STB";
        const DANH_SACH_MA = [ma1, ma2];
        const URL = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(",")}`;
        const response = await axios.get(URL);
        const data = response.data;
        expect(data[0].sym).toBe(ma1);
        expect(data[1].sym).toBe(ma2);
    });
});

describe("kiểm tra url cafef chạy chính xác", () => {
    test("kiểm tra phản hồi từ api báo cáo phân tích", async () => {
        const tenMa = "HPG";
        const URL = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
        const response = await axios.post(URL);
        const data = response.data.Data;
        expect(data[0].Title).toContain(tenMa);
    });

    test("kiểm tra phản hồi từ api cafef", async () => {
        const tenMa = "HPG";
        const URL = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
        const response = await axios.get(URL);
        const data = response.data;
        expect(data).toContain('id="divEvents"');
    });

    test("kiểm tra phản hồi từ api cafef báo cáo tài chính", async () => {
        const tenMa = "HPG";
        const URL = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
        const response = await axios.get(URL);
        const data = response.data;
        expect(data).toContain('Báo cáo tài chính');
    });
});

describe("kiểm tra url tcbs chạy chính xác", () => {
    test("kiểm tra phản hồi từ api vps", async () => {
        const tenMa = "HPG";
        const URL = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
        const response = await axios.get(URL);
        const data = response.data.listShareHolder;
        expect(data[0].ticker).toBe(tenMa);
    });
});