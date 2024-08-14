/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { DateHelper } from '../utility/DateHelper';
import { HttpHelper } from '../utility/HttpHelper';
import { LogHelper } from '../utility/LogHelper';
import { SheetHelper, SHEET_NAME } from '../utility/SheetHelper';
import { ResponseSsi, ResponseVndirect } from '../types/types';

function layChiSoVnIndex(): void {
  const ngayHienTai: string = DateHelper.layNgayHienTai('YYYY-MM-DD');
  const duLieuNgayMoiNhat: string = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'A1');
  const thanhKhoanMoiNhat: number = parseFloat(new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'D1'));
  const url = 'https://banggia.cafef.vn/stockhandler.ashx?index=true';

  const object = new HttpHelper().sendPostRequest(url);
  const duLieuNhanVe = object[1];
  const thanhKhoan: number = parseFloat(duLieuNhanVe.value.replace(/,/g, '')) * 1000000000;
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SHEET_NAME.sheetHose, 'A1:D1');
  } else if (thanhKhoan !== thanhKhoanMoiNhat) {
    new SheetHelper().chen1HangVaoDauSheet(SHEET_NAME.sheetHose);
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SHEET_NAME.sheetHose, 'A1:D1');
  } else {
    console.log('No action required');
  }
}

function layThongTinCoBan(): void {
  const danhSachMa: string[] = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  layThongTinPB(danhSachMa);
  layThongTinPE(danhSachMa);
  layThongTinRoomNuocNgoai(danhSachMa);
  layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa);
}

// Hàm lấy giá, khối lượng và thông tin mua bán của khối ngoại hàng ngày
function layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay(): void {
  layGiaHangNgay();
  layKhoiLuongHangNgay();
  layKhoiNgoaiMuaHangNgay();
  layKhoiNgoaiBanHangNgay();
  LogHelper.logTime(SHEET_NAME.sheetCauHinh, 'G4');
}

function layGiaThamChieu(): void {
  const DEFAULT_FORMAT = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B6');
  const DANH_SACH_MA: string[] = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  const date: string = DateHelper.doiDinhDangNgay(new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B1'), DEFAULT_FORMAT, 'DD/MM/YYYY');
  const market = 'HOSE';
  let index = 2;
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${date}&lookupRequest.toDate=${date}&lookupRequest.market=${market}`;
  const token = new HttpHelper().getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = new HttpHelper().sendRequest(URL, OPTION);
  const datas = object.data;
  for (const e of datas) {
    const vitri = new SheetHelper().layViTriCotThamChieu(e.Symbol, DANH_SACH_MA, 2);
    if (e.Symbol.length === 3 && vitri !== -1) {
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[e.ClosePrice]], SHEET_NAME.sheetDuLieu, `C${vitri}:C${vitri}`);
      index++;
    }
  }
}

function layThongTinPB(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const URL = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object: any = new HttpHelper().sendGetRequest(URL);

    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = new SheetHelper().layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[value]], SHEET_NAME.sheetDuLieu, `E${vitri}:E${vitri}`);
    });
  }
}

function layThongTinPE(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = new HttpHelper().sendGetRequest(url);
    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = new SheetHelper().layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[value]], SHEET_NAME.sheetDuLieu, `F${vitri}:F${vitri}`);
    });
  }
}

function layThongTinRoomNuocNgoai(danhSachMa: string[]): void {
  const QUERY_API = 'https://finfo-api.vndirect.com.vn/v4';
  const duLieuCotThamChieu = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = new HttpHelper().sendGetRequest(url);
    object.data.forEach((element: { code?: string; totalRoom?: number; currentRoom?: number }) => {
      const totalRoom: number = element.totalRoom ?? 0;
      const currentRoom: number = element.currentRoom ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = new SheetHelper().layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[totalRoom, currentRoom]], SHEET_NAME.sheetDuLieu, `G${vitri}:H${vitri}`);
    });
  }
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = new HttpHelper().sendGetRequest(url);
    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = new SheetHelper().layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[value]], SHEET_NAME.sheetDuLieu, `I${vitri}:I${vitri}`);
    });
  }
}

export function duLieuTam(): void {
  new SheetHelper().layDuLieuTrongCot('TRUY VAN', 'A').forEach((date: string) => {
    // layKhoiNgoaiBanHangNgay('Tam', date);
    layKhoiNgoaiMuaHangNgay('Tam', date);
    // layKhoiLuongHangNgay('Tam', date);
    // layGiaHangNgay('Tam', date);
  });
}

export function layKhoiNgoaiBanHangNgay(sheetName = SHEET_NAME.sheetKhoiNgoaiBan, date = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'A1')) {
  const danhSachMa = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  const hangCuoi = new SheetHelper().laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = new SheetHelper().layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = new HttpHelper().sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = new SheetHelper().layDuLieuTrongHang(sheetName, 1);
        new SheetHelper().ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: ResponseVndirect) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              new SheetHelper().ghiDuLieuVaoDay([[item.sellVol]], sheetName, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối ngoại bán hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}

function layKhoiNgoaiMuaHangNgay(sheetName = SHEET_NAME.sheetKhoiNgoaiMua, date = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'A1')) {
  const danhSachMa = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  const hangCuoi = new SheetHelper().laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = new SheetHelper().layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = new HttpHelper().sendGetRequest(url);
      if (object?.data.length > 0) {
        const header = new SheetHelper().layDuLieuTrongHang(sheetName, 1);
        new SheetHelper().ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: ResponseVndirect) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              new SheetHelper().ghiDuLieuVaoDay([[item.buyVol]], sheetName, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối ngoại mua hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}

function layKhoiLuongHangNgay(sheetName = SHEET_NAME.sheetKhoiLuong, date = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'A1')) {
  const danhSachMa = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  const hangCuoi = new SheetHelper().laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = new SheetHelper().layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(',')}~date:gte:${date}~date:lte:${date}`;
      const object = new HttpHelper().sendGetRequest(url);
      if (object?.data.length > 0) {
        const header = new SheetHelper().layDuLieuTrongHang(sheetName, 1);
        new SheetHelper().ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: ResponseVndirect) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              new SheetHelper().ghiDuLieuVaoDay([[item.nmVolume]], sheetName, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối lượng hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}
function layGiaHangNgay(sheetName = SHEET_NAME.sheetGia, date = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetHose, 'A1')) {
  const DEFAULT_FORMAT = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B6');
  const fromDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'DD/MM/YYYY');
  const toDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'DD/MM/YYYY');
  const hangCuoi = new SheetHelper().laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = new SheetHelper().layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  const market = 'HOSE';
  if (duLieuNgayMoiNhat !== date) {
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    const token = new HttpHelper().getToken();
    const OPTION: URLFetchRequestOptions = {
      method: 'get',
      // eslint-disable-next-line @typescript-eslint/naming-convention
      headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
    };
    const object = new HttpHelper().sendRequest(url, OPTION);
    if (object?.data.length > 0) {
      const header = new SheetHelper().layDuLieuTrongHang(sheetName, 1);
      new SheetHelper().ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: ResponseSsi) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.Symbol) {
            new SheetHelper().ghiDuLieuVaoDay([[item.ClosePrice]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy giá hàng ngày thành công');
  } else {
    console.log('done');
  }
}

function layDanhSachMa(): void {
  const market = 'HOSE';
  const pageIndex = 1;
  const pageSize = 1000;
  const token = new HttpHelper().getToken();
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const mangDuLieuChinh: Array<[string, string]> = [];
  const response = new HttpHelper().sendRequest(url, OPTION);
  const datas = response.data;
  new SheetHelper().xoaDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A', 2, 2);
  for (const element of datas) {
    if (element.Symbol.length == 3) {
      mangDuLieuChinh.push([element.Symbol, element.StockName]);
    }
  }
  new SheetHelper().ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SHEET_NAME.sheetDuLieu, 2, 'A');
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_DANH_MUC_DC(url: string) {
  const result: any = [];
  const response = new HttpHelper().sendRequest(url);
  const data = response.ffs_holding;
  // eslint-disable-next-line @typescript-eslint/naming-convention
  data.forEach((element: { stock?: string; sector_vi?: string; per_nav?: string; bourse_en?: string; created?: string }) => {
    const tenMa = element.stock ?? '_';
    const nhomNganh = element.sector_vi ?? '_';
    const tyLe = element.per_nav ?? '_';
    const sanGD = element.bourse_en ?? '_';
    const capNhatLuc = element.created ?? '-';
    result.push([tenMa, nhomNganh, sanGD, tyLe, capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_TAI_SAN_DC(url: string) {
  const result: any = [];
  const response = new HttpHelper().sendRequest(url);
  const data = response.ffs_asset;
  console.log(response);
  // eslint-disable-next-line @typescript-eslint/naming-convention
  data.forEach((element: { name_vi?: string; weight?: string; created?: string }) => {
    const tenTaiSan = element.name_vi ?? '_';
    const tyLe = element.weight ?? '_';
    const capNhatLuc = element.created ?? '-';
    result.push([tenTaiSan, tyLe, capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_SU_KIEN() {
  const result: any = [];
  const content = UrlFetchApp.fetch(`https://hontrang.github.io/tradingeconomics/`).getContentText();
  const $ = Cheerio.load(content);
  let date: string;
  $('table#calendar>thead.table-header,table#calendar>tbody').each(function (this: any) {
    if ($(this).attr('class') !== undefined) {
      date = DateHelper.doiDinhDangNgay($(this).find('tr>th:nth-child(1)').text().trim(), 'dddd MMMM DD YYYY', 'ddd YYYY/MM/DD');
    } else {
      $(this)
        .children('tr')
        .each(function (this: any) {
          const timeValue = $(this).find('td:nth-child(1)').text().trim() || '00:00 AM';
          const thoigian = `${date} ${timeValue}`;
          const currency = $(this).find('td:nth-child(2) td.calendar-iso').text().trim();
          const name = $(this).find('td:nth-child(3)').text().trim();
          const actual = $(this).find('td:nth-child(4)').text().trim();
          const previous = $(this).find('td:nth-child(5)').text().trim();
          const forecast = $(this).find('td:nth-child(6)').text().trim();
          result.push([thoigian, currency, name, actual, forecast, previous]);
        });
    }
  });
  return result;
}
