/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { DateHelper } from '@utils/DateHelper';
import { HttpHelper } from '@utils/HttpHelper';
import { LogHelper } from '@utils/LogHelper';
import { SheetHelper } from '@utils/SheetHelper';
import { ResponseDC, ResponseSsi, ResponseVndirect } from '@src/types/types';

function layChiSoVnIndex(): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const ngayHienTai: string = DateHelper.layNgayHienTai('YYYY-MM-DD');
  const duLieuNgayMoiNhat: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1');
  const thanhKhoanMoiNhat: number = parseFloat(sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'D1'));
  const url = 'https://banggia.cafef.vn/stockhandler.ashx?index=true';

  const object = httpHelper.sendPostRequest(url);
  const duLieuNhanVe = object[1];
  const thanhKhoan: number = parseFloat(duLieuNhanVe.value.replace(/,/g, '')) * 1000000000;
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    sheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SheetHelper.sheetName.sheetHose, 'A1:D1');
  } else if (thanhKhoan !== thanhKhoanMoiNhat) {
    sheetHelper.chen1HangVaoDauSheet(SheetHelper.sheetName.sheetHose);
    sheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SheetHelper.sheetName.sheetHose, 'A1:D1');
  } else {
    console.log('No action required');
  }
}

function layThongTinCoBan(): void {
  const sheetHelper = new SheetHelper();
  const danhSachMa: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');
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
  LogHelper.logTime(SheetHelper.sheetName.sheetCauHinh, 'G4');
}

async function layGiaThamChieu(): Promise<void> {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const DEFAULT_FORMAT = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const DANH_SACH_MA: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');
  const date: string = DateHelper.doiDinhDangNgay(sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B1'), DEFAULT_FORMAT, 'dd/MM/yyyy');
  const market = 'HOSE';
  let index = 2;
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${date}&lookupRequest.toDate=${date}&lookupRequest.market=${market}`;
  const token = await httpHelper.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = httpHelper.sendRequest(URL, OPTION);
  const datas = object.data;
  for (const e of datas) {
    const vitri = sheetHelper.layViTriCotThamChieu(e.Symbol, DANH_SACH_MA, 2);
    if (e.Symbol.length === 3 && vitri !== -1) {
      sheetHelper.ghiDuLieuVaoDayTheoVung([[e.ClosePrice]], SheetHelper.sheetName.sheetDuLieu, `C${vitri}:C${vitri}`);
      index++;
    }
  }
}

function layThongTinPB(danhSachMa: string[]): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const URL = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(URL);

    object.data.forEach((element: ResponseVndirect) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      sheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.sheetName.sheetDuLieu, `E${vitri}:E${vitri}`);
    });
  }
}

function layThongTinPE(danhSachMa: string[]): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const URL = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(URL);
    object.data.forEach((element: ResponseVndirect) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      sheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.sheetName.sheetDuLieu, `F${vitri}:F${vitri}`);
    });
  }
}

function layThongTinRoomNuocNgoai(danhSachMa: string[]): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const QUERY_API = 'https://finfo-api.vndirect.com.vn/v4';
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const URL = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(URL);
    object.data.forEach((element: ResponseVndirect) => {
      const totalRoom: number = element.totalRoom ?? 0;
      const currentRoom: number = element.currentRoom ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      sheetHelper.ghiDuLieuVaoDayTheoVung([[totalRoom, currentRoom]], SheetHelper.sheetName.sheetDuLieu, `G${vitri}:H${vitri}`);
    });
  }
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa: string[]): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(url);
    object.data.forEach((element: ResponseVndirect) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      sheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.sheetName.sheetDuLieu, `I${vitri}:I${vitri}`);
    });
  }
}

export function duLieuTam(): void {
  const sheetHelper = new SheetHelper();
  sheetHelper.layDuLieuTrongCot('TRUY VAN', 'A').forEach((date: string) => {
    layKhoiNgoaiBanHangNgay('Tam', date);
    layKhoiNgoaiMuaHangNgay('Tam', date);
    layKhoiLuongHangNgay('Tam', date);
    layGiaHangNgay('Tam', date);
  });
}

export function layKhoiNgoaiBanHangNgay(sheetName = SheetHelper.sheetName.sheetKhoiNgoaiBan, date = new SheetHelper().layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1')) {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiNgoaiBan, 1).slice(1);
  const hangCuoi = sheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = sheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${headers.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
    const object = httpHelper.sendGetRequest(URL);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: ResponseVndirect) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.code) {
            sheetHelper.ghiDuLieuVaoDay([[item.sellVol]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy khối ngoại bán hàng ngày thành công');
  } else {
    console.log('done');
  }
}

function layKhoiNgoaiMuaHangNgay(sheetName = SheetHelper.sheetName.sheetKhoiNgoaiMua, date = new SheetHelper().layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1')) {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiNgoaiMua, 1).slice(1);
  const hangCuoi = sheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = sheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    const url = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${headers.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
    const object = httpHelper.sendGetRequest(url);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: ResponseVndirect) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.code) {
            sheetHelper.ghiDuLieuVaoDay([[item.buyVol]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy khối ngoại mua hàng ngày thành công');
  } else {
    console.log('done');
  }
}

function layKhoiLuongHangNgay(sheetName = SheetHelper.sheetName.sheetKhoiLuong, date = new SheetHelper().layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1')) {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiLuong, 1).slice(1);
  const hangCuoi = sheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = sheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${headers.join(',')}~date:gte:${date}~date:lte:${date}`;
    const object = httpHelper.sendGetRequest(URL);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: ResponseVndirect) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.code) {
            sheetHelper.ghiDuLieuVaoDay([[item.nmVolume]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy khối lượng hàng ngày thành công');
  } else {
    console.log('done');
  }
}
async function layGiaHangNgay(sheetName = SheetHelper.sheetName.sheetGia, date = new SheetHelper().layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1')) {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const DEFAULT_FORMAT = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const fromDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'dd/MM/yyyy');
  const toDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'dd/MM/yyyy');
  const hangCuoi = sheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = sheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  const market = 'HOSE';
  if (duLieuNgayMoiNhat !== date) {
    const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    const token = await httpHelper.getToken();
    const OPTION: URLFetchRequestOptions = {
      method: 'get',
      // eslint-disable-next-line @typescript-eslint/naming-convention
      headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
    };
    const object = httpHelper.sendRequest(URL, OPTION);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: ResponseSsi) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.Symbol) {
            sheetHelper.ghiDuLieuVaoDay([[item.ClosePrice]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy giá hàng ngày thành công');
  } else {
    console.log('done');
  }
}

async function layDanhSachMa(): Promise<void> {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const market = 'HOSE';
  const pageIndex = 1;
  const pageSize = 1000;
  const token = await httpHelper.getToken();
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const mangDuLieuChinh: Array<[string, string]> = [];
  const response = httpHelper.sendRequest(URL, OPTION);
  const datas = response.data;
  sheetHelper.xoaDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A', 2, 2);
  for (const element of datas) {
    if (element.Symbol.length == 3) {
      mangDuLieuChinh.push([element.Symbol, element.StockName]);
    }
  }
  sheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.sheetName.sheetDuLieu, 2, 'A');
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_DANH_MUC_DC(URL: string) {
  const httpHelper = new HttpHelper();
  const result: string[][] = [];
  const response = httpHelper.sendRequest(URL);
  const data = response.ffs_holding;
  // eslint-disable-next-line @typescript-eslint/naming-convention
  data.forEach((element: ResponseDC) => {
    const tenMa = element.stock ?? '_';
    const nhomNganh = element.sector_vi ?? '_';
    const tyLe = element.per_nav ?? 0;
    const sanGD = element.bourse_en ?? '_';
    const capNhatLuc = element.created ?? '-';
    result.push([tenMa, nhomNganh, sanGD, String(tyLe), capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_TAI_SAN_DC(URL: string) {
  const httpHelper = new HttpHelper();
  const result: string[][] = [];
  const response = httpHelper.sendRequest(URL);
  const data = response.ffs_asset;
  console.log(response);
  // eslint-disable-next-line @typescript-eslint/naming-convention
  data.forEach((element: ResponseDC) => {
    const tenTaiSan = element.name_vi ?? '_';
    const tyLe = element.weight ?? '_';
    const capNhatLuc = element.created ?? '-';
    result.push([tenTaiSan, String(tyLe), capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_SU_KIEN() {
  const result: string[][] = [];
  const content = UrlFetchApp.fetch(`https://hontrang.github.io/google-sheet/`).getContentText();
  const $ = Cheerio.load(content);
  let date: string;
  $('table#calendar>thead.table-header,table#calendar>tbody').each(function () {
    if ($(this).attr('class') !== undefined) {
      date = DateHelper.doiDinhDangNgay($(this).find('tr>th:nth-child(1)').text().trim(), 'DDDD MMMM dd yyyy', 'DDD yyyy/MM/dd');
    } else {
      $(this)
        .children('tr')
        .each(function () {
          const timeValue = $(this).find('td:nth-child(1)').text().trim() || '00:00 AM';
          const thoigian = `${date} ${timeValue}`;
          const currency = $(this).find('td:nth-child(2) td.calendar-iso').text().trim();
          const name = $(this).find('td:nth-child(3)').text().trim();
          const actual = $(this).find('td:nth-child(4)').text().trim();
          const previous = $(this).find('td:nth-child(5)').text().trim();
          const consensus = $(this).find('td:nth-child(6)').text().trim();
          const forecast = $(this).find('td:nth-child(7)').text().trim();
          result.push([thoigian, currency, name, actual, previous, consensus, forecast]);
        });
    }
  });
  return result;
}
