/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { DateHelper } from '../utility/DateHelper';
import { HttpHelper } from '../utility/HttpHelper';
import { LogHelper } from '../utility/LogHelper';
import { SheetHelper } from '../utility/SheetHelper';

function layChiSoVnIndex(): void {
  const ngayHienTai: string = DateHelper.layNgayHienTai('YYYY-MM-DD');
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'A1');
  const thanhKhoanMoiNhat: number = parseFloat(SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'D1'));
  const url = 'https://banggia.cafef.vn/stockhandler.ashx?index=true';

  const object = HttpHelper.sendPostRequest(url);
  const duLieuNhanVe = object[1];
  const thanhKhoan: number = parseFloat(duLieuNhanVe.value.replace(/,/g, '')) * 1000000000;
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    SheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SheetHelper.SheetName.SHEET_HOSE, 'A1:D1');
  } else if (thanhKhoan !== thanhKhoanMoiNhat) {
    SheetHelper.chen1HangVaoDauSheet(SheetHelper.SheetName.SHEET_HOSE);
    SheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]], SheetHelper.SheetName.SHEET_HOSE, 'A1:D1');
  } else {
    console.log('No action required');
  }
}

function layThongTinCoBan(): void {
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');
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
  LogHelper.logTime(SheetHelper.SheetName.SHEET_CAU_HINH, 'G4');
}

function layGiaThamChieu(): void {
  const DEFAULT_FORMAT = SheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B6');
  const DANH_SACH_MA: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');
  const date: string = DateHelper.doiDinhDangNgay(SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_CAU_HINH, 'B1'), DEFAULT_FORMAT, 'DD/MM/YYYY');
  const market = 'HOSE';
  let index = 2;
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${date}&lookupRequest.toDate=${date}&lookupRequest.market=${market}`;
  const token = HttpHelper.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = HttpHelper.sendRequest(URL, OPTION);
  const datas = object.data;
  for (const e of datas) {
    const vitri = SheetHelper.layViTriCotThamChieu(e.Symbol, DANH_SACH_MA, 2);
    if (e.Symbol.length === 3 && vitri !== -1) {
      SheetHelper.ghiDuLieuVaoDayTheoVung([[e.ClosePrice]], SheetHelper.SheetName.SHEET_DU_LIEU, `C${vitri}:C${vitri}`);
      index++;
    }
  }
}

function layThongTinPB(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const URL = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object: any = HttpHelper.sendGetRequest(URL);

    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SheetName.SHEET_DU_LIEU, `E${vitri}:E${vitri}`);
    });
  }
}

function layThongTinPE(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object = HttpHelper.sendGetRequest(url);
    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SheetName.SHEET_DU_LIEU, `F${vitri}:F${vitri}`);
    });
  }
}

function layThongTinRoomNuocNgoai(danhSachMa: string[]): void {
  const QUERY_API = 'https://finfo-api.vndirect.com.vn/v4';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object = HttpHelper.sendGetRequest(url);
    object.data.forEach((element: { code?: string; totalRoom?: number; currentRoom?: number }) => {
      const totalRoom: number = element.totalRoom ?? 0;
      const currentRoom: number = element.currentRoom ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[totalRoom, currentRoom]], SheetHelper.SheetName.SHEET_DU_LIEU, `G${vitri}:H${vitri}`);
    });
  }
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object = HttpHelper.sendGetRequest(url);
    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SheetName.SHEET_DU_LIEU, `I${vitri}:I${vitri}`);
    });
  }
}

function duLieuTam(): void {
  SheetHelper.layDuLieuTrongCot('TRUY VAN', 'A').forEach((date: string) => {
    layKhoiNgoaiBanHangNgay('Tam', date);
    layKhoiNgoaiMuaHangNgay('Tam', date);
    layKhoiLuongHangNgay('Tam', date);
    layGiaHangNgay('Tam', date);
  });
}

function layKhoiNgoaiBanHangNgay(sheetName = SheetHelper.SheetName.SHEET_KHOI_NGOAI_BAN, date = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'A1')) {
  const danhSachMa = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');
  const hangCuoi = SheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = SheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = HttpHelper.sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = SheetHelper.layDuLieuTrongHang(sheetName, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.sellVol]], sheetName, hangCuoi + 1, i + 1);
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

function layKhoiNgoaiMuaHangNgay(sheetName = SheetHelper.SheetName.SHEET_KHOI_NGOAI_MUA, date = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'A1')) {
  const danhSachMa = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');
  const hangCuoi = SheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = SheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = HttpHelper.sendGetRequest(url);
      if (object?.data.length > 0) {
        const header = SheetHelper.layDuLieuTrongHang(sheetName, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.buyVol]], sheetName, hangCuoi + 1, i + 1);
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

function layKhoiLuongHangNgay(sheetName = SheetHelper.SheetName.SHEET_KHOI_LUONG, date = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'A1')) {
  const danhSachMa = SheetHelper.layDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A');
  const hangCuoi = SheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = SheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(',')}~date:gte:${date}~date:lte:${date}`;
      const object = HttpHelper.sendGetRequest(url);
      if (object?.data.length > 0) {
        const header = SheetHelper.layDuLieuTrongHang(sheetName, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.nmVolume]], sheetName, hangCuoi + 1, i + 1);
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
function layGiaHangNgay(sheetName = SheetHelper.SheetName.SHEET_GIA, date = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SheetName.SHEET_HOSE, 'A1')) {
  const DEFAULT_FORMAT = SheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B6');
  const fromDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'DD/MM/YYYY');
  const toDate = DateHelper.doiDinhDangNgay(date, DEFAULT_FORMAT, 'DD/MM/YYYY');
  const hangCuoi = SheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = SheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  const market = 'HOSE';
  if (duLieuNgayMoiNhat !== date) {
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    const token = HttpHelper.getToken();
    const OPTION: URLFetchRequestOptions = {
      method: 'get',
      // eslint-disable-next-line @typescript-eslint/naming-convention
      headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
    };
    const object = HttpHelper.sendRequest(url, OPTION);
    if (object?.data.length > 0) {
      const header = SheetHelper.layDuLieuTrongHang(sheetName, 1);
      SheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: any) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.Symbol) {
            SheetHelper.ghiDuLieuVaoDay([[item.ClosePrice]], sheetName, hangCuoi + 1, i + 1);
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
  const token = HttpHelper.getToken();
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const mangDuLieuChinh: Array<[string, string]> = [];
  const response = HttpHelper.sendRequest(url, OPTION);
  const datas = response.data;
  SheetHelper.xoaDuLieuTrongCot(SheetHelper.SheetName.SHEET_DU_LIEU, 'A', 2, 2);
  for (const element of datas) {
    if (element.Symbol.length == 3) {
      mangDuLieuChinh.push([element.Symbol, element.StockName]);
    }
  }
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SheetName.SHEET_DU_LIEU, 2, 'A');
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_DANH_MUC_DC(url: string) {
  const result: any = [];
  const response = HttpHelper.sendRequest(url);
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
  const response = HttpHelper.sendRequest(url);
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
