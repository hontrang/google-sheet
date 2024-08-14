import { ExcelHelper } from "./ExcelHelper";
import { AxiosHelper } from "./AxiosHelper";
import { SheetHelper } from "../utility/SheetHelper";
import { ResponseVndirect } from "../types/types";

const sheetHelper = new ExcelHelper();
const httpHelper = new AxiosHelper();

async function main() {
    await sheetHelper.truyCapWorkBook();
    const danhSachMa: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');
    await layThongTinPE(danhSachMa);
    await sheetHelper.luuWorkBook();
}

async function layThongTinPE(danhSachMa: string[]): Promise<void> {
    const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
    const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');

    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.join(',')}`;
    const object = await httpHelper.sendGetRequest(url);
    object.data.data.forEach((element: ResponseVndirect) => {
        const value: number = element.value ?? 0;
        const tenMa: string = element.code ?? '_';
        const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
        sheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.sheetName.sheetDuLieu, `F${vitri}:F${vitri}`);
    });
}

main();

