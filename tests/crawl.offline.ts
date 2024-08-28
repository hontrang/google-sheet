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
  try {
    for (const element of headers) {
      const tenMa = element;
      await layThongTin(tenMa, headers, token);
    }
  } catch (error) {
    await sheetHelper.luuWorkBook();
  }
  await sheetHelper.luuWorkBook();
});

async function layThongTin(tenMa: string, headers: string[], token: string): Promise<void> {
  const pageIndex = 1;
  const pageSize = 1000;

  const toDate = '01/01/2013';
  const ascending = true;
  const hangCuoiCungTrongSheet = sheetHelper.laySoHangTrongSheet(tenMa);

  const fromDate = sheetHelper.layDuLieuTrongO(tenMa, `A${hangCuoiCungTrongSheet}`);
  const year = fromDate.split(`/`)[2]
  if (hangCuoiCungTrongSheet < 2 || Number(year) >= 2013 || fromDate == '28/12/2012') {
    console.log(`skip ${tenMa}`);
    return;
  };


  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
  await LogHelper.sleepSync(1000);
  const response = await httpHelper.sendRequest(URL, { headers: { Authorization: token } });
  sheetHelper.taoSheetMoi(tenMa);
  sheetHelper.ghiDuLieuVaoO(tenMa, tenMa, 'B1');
  console.log(response.data.message);
  const datas = response.data.data;
  datas.forEach(function (element: ResponseSsi, index: number) {
    sheetHelper.ghiDuLieuVaoO(element.TradingDate, tenMa, `A${index + 1 + hangCuoiCungTrongSheet}`);
    sheetHelper.ghiDuLieuVaoO(`${element.Close}`, tenMa, `B${index + 1 + hangCuoiCungTrongSheet}`);
  });
  console.log(`Lấy thông tin mã ${tenMa} thành công`);
}
