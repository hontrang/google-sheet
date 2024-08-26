import { test } from '@playwright/test';
import { ExcelHelper } from '@src/offline/ExcelHelper';
import { AxiosHelper } from '@src/offline/AxiosHelper';
import { SheetHelper } from '@src/utility/SheetHelper';
import { ResponseSsi, ResponseVndirect } from '@src/types/types';
import { LogHelper } from '@utils/LogHelper';

const sheetHelper = new ExcelHelper();
const httpHelper = new AxiosHelper();

test('crawl data', async () => {
  await sheetHelper.truyCapWorkBook();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiNgoaiMua, 1).slice(1);
  sheetHelper.ghiDuLieuVaoDay(headers, 'Tam', 1, 2);
  // for (let i = 0; i < headers.length; i++) {
  // const tenMa = headers[i];
  await layThongTin("AAA");
  // }
  await sheetHelper.luuWorkBook();
});

async function layThongTin(tenMa: string): Promise<void> {
  const pageIndex = 1;
  const pageSize = 1000;
  const fromDate = '01/01/2000';
  const toDate = '04/05/2020';
  const ascending = true;
  const token = await httpHelper.getToken();
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyOhlc?lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.ascending=${ascending}&lookupRequest.Symbol=${tenMa}`;
  const response = await httpHelper.sendRequest(URL, { headers: { Authorization: token } });
  LogHelper.sleep(100);
  const datas = response.data.data;
  let index = 2;
  datas.forEach(function (element: ResponseSsi) {
    sheetHelper.ghiDuLieuVaoO(element.TradingDate, 'Tam', `A${index}`);
    index++;
  });

  // const khoiLuong: string[] = [date];
  // if (datas.length > 0) {
  //   for (const head of headers) {
  //     const element: ResponseVndirect = timDoiTuongCoMa(datas, head);
  //     if (element.nmVolume === undefined) {
  //       throw new Error(`Sai du lieu`);
  //     } else {
  //       khoiLuong.push(String(element.nmVolume));
  //     }
  //   }
  // }
  // sheetHelper.ghiDuLieuVaoDay(khoiLuong, 'Tam', hang, 1);
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
