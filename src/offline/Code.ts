import { ExcelHelper } from "../offline/utility/ExcelHelper";
import { AxiosHelper } from "../offline/utility/AxiosHelper";
import { ResponseVndirect } from "../types/types";
import { sleepSync } from "bun";

const excel = new ExcelHelper();
const axios = new AxiosHelper();

async function main() {

    await excel.truyCapWorkBook();
    const ngay = excel.layDuLieuTrongCot('TRUY VAN', 'A').reverse();
    for (const date of ngay) {
        await layKhoiNgoaiMuaHangNgay('Tam', date);
    }
    await excel.luuWorkBook();
}

async function layKhoiNgoaiMuaHangNgay(sheetName = 'Tam', date = '2024-01-01'): Promise<void> {
    // const headers = excel.layDuLieuTrongHang(sheetName, 1).slice(1);
    const tenMa = 'VIC';
    const url = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${tenMa}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
    try {
        const response = await axios.sendGetRequest(url);
        if (response.data.length > 0) {
            // const list = [date];
            // response.data.map(async (item: ResponseVndirect) => {
            //     for (const header of headers) {
            //         if (header === item.code) {
            //             list.push(item.buyVol);
            //         }
            //     }
            // });
            const item: ResponseVndirect = response.data[0];
            const hangCuoi = await excel.laySoHangTrongSheet(sheetName);
            await excel.ghiDuLieuVaoDay([date, item.sellVol], sheetName, hangCuoi + 1, 1);
            console.log('Lấy khối ngoại bán hàng ngày thành công');
        } else {
            const hangCuoi = await excel.laySoHangTrongSheet(sheetName);
            await excel.ghiDuLieuVaoDay([`${date}`, ""], sheetName, hangCuoi + 1, 1);
        }
        sleepSync(100);
    } catch (error) {
        console.error('Lỗi khi lấy dữ liệu từ API:', error);
    }
}

main();

