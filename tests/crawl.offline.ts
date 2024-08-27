import { test } from '@playwright/test';
import { ExcelHelper } from '@src/offline/ExcelHelper';
import { AxiosHelper } from '@src/offline/AxiosHelper';
import { SheetHelper } from '@src/utility/SheetHelper';
import { ResponseSsi } from '@src/types/types';
import { LogHelper } from '@utils/LogHelper';

const sheetHelper = new ExcelHelper();
const httpHelper = new AxiosHelper();

test('crawl data', async () => {
  await sheetHelper.truyCapWorkBook();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetGia, 1).slice(1);
  const token = await httpHelper.getToken();
  let ghiNgay = true;
  for (const element of headers) {
    const tenMa = element;
    await layThongTin(tenMa, headers, token);
  }
  await sheetHelper.luuWorkBook();
});

async function layThongTin(tenMa: string, headers: string[], token: string): Promise<void> {
  const pageIndex = 1;
  const pageSize = 1000;
  const fromDate = '01/01/2000';
  const toDate = '04/05/2020';
  const ascending = true;
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
  const response = await httpHelper.sendRequest(URL, { headers: { Authorization: token } });
  LogHelper.sleep(100);
  sheetHelper.taoSheetMoi(tenMa);
  sheetHelper.ghiDuLieuVaoO(tenMa, tenMa, 'B1');
  const datas = response.data.data;
  datas.forEach(function (element: ResponseSsi, index: number) {
    sheetHelper.ghiDuLieuVaoO(element.TradingDate, tenMa, `A${index + 2}`);
    sheetHelper.ghiDuLieuVaoO(`${element.Close}`, tenMa, `B${index + 2}`);
  });
  console.log(`Lấy thông tin mã ${tenMa} thành công`);
}
