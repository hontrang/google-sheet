import { ExcelHelper } from "@src/offline/ExcelHelper";
import { AxiosHelper } from "@src/offline/AxiosHelper";
import { SheetHelper } from "@src/utility/SheetHelper";
import { ResponseVndirect } from "@src/types/types";

const sheetHelper = new ExcelHelper();
const httpHelper = new AxiosHelper();

async function main() {
    await sheetHelper.truyCapWorkBook();
    const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiNgoaiMua, 1).slice(1);
    const dates = sheetHelper.layDuLieuTrongCot('TRUY VAN', 'A');
    for (let i = 0; i < dates.length; i++) {
        const date = dates[i];
        await layThongTin(headers, date, i + 2);
    }
    await sheetHelper.luuWorkBook();
}

async function layThongTin(headers: string[], date: string, hang: number): Promise<void> {
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${headers.join(',')}~date:gte:${date}~date:lte:${date}`;
    const response = await httpHelper.sendGetRequest(URL);
    await Bun.sleepSync(100);
    const datas = response.data.data;
    const khoiLuong: any[] = [date];
    if (datas.length > 0) {
        for (const head of headers) {
            const element: ResponseVndirect = timDoiTuongCoMa(datas, head);
            if (element.nmVolume === undefined) {
                throw new Error(`Sai du lieu`);
            } else {
                khoiLuong.push(element.nmVolume);
            }

        }
    }
    sheetHelper.ghiDuLieuVaoDay(khoiLuong, "Tam", hang, 1);
    console.log('Lấy thông tin thành công');
}

function timDoiTuongCoMa(danhSach: ResponseVndirect[], tenMa: string): ResponseVndirect {
    const result = danhSach.find((element) => element.code === tenMa);
    if (result === undefined) {
        return {
            buyVol: 0,
            sellVol: 0,
            nmVolume: 0
        };
    }
    return result;
}

main();

