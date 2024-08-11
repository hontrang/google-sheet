import { ExcelHelper } from "../utility/ExcelHelper";
import { SheetHelper } from "../utility/SheetHelper";
import { AxiosHelper } from "../utility/AxiosHelper";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;


async function main() {
    const excel = new ExcelHelper();
    const axiosHelper = new AxiosHelper();
    await excel.truyCapWorkBook();
    const market = 'HOSE';
    const pageIndex = 1;
    const pageSize = 1000;
    const token = await axiosHelper.getToken();
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
    const OPTION: URLFetchRequestOptions = {
        method: 'get',
        // eslint-disable-next-line @typescript-eslint/naming-convention
        headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
    };
    const mangDuLieuChinh: Array<[string, string]> = [];
    const response = axiosHelper.sendRequest(url, OPTION);
    const datas = response.data;
    excel.xoaDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A', 2, 2);
    for (const element of datas) {
        if (element.Symbol.length == 3) {
            mangDuLieuChinh.push([element.Symbol, element.StockName]);
        }
    }
    excel.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SheetName.SHEET_DU_LIEU, 2, 'A');
    await excel.luuWorkBook();
}

main();

