/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import { DateHelper } from '@utils/DateHelper';
import { HttpHelper } from '@utils/HttpHelper';
import { LogHelper } from '@utils/LogHelper';
import { SheetHelper } from '@utils/SheetHelper';
import { IResponseDC, IResponseSimplize, IResponseVndirect } from '@src/types/generic';

function layChiSoVnIndex(): void {

  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const ngayHienTai: string = DateHelper.layNgayHienTai('yyyy-MM-dd');
  const duLieuNgayMoiNhat: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1');
  const thanhKhoanMoiNhat: number = parseFloat(sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'D1'));
  const tenMa = 'VNINDEX';
  const url = `https://api2.simplize.vn/api/historical/quote/${tenMa}?type=index`;

  const response = httpHelper.sendGetRequest(url);
  const duLieuNhanVe: IResponseSimplize = response.data;
  const thanhKhoan: number = duLieuNhanVe.totalValue ?? 0;
  const tiLeThayDoi: number = duLieuNhanVe.pctChange ?? 0;
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    sheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe?.priceClose, tiLeThayDoi / 100, thanhKhoan]], SheetHelper.sheetName.sheetHose, 'A1:D1');
  } else if (thanhKhoan !== thanhKhoanMoiNhat) {
    sheetHelper.chen1HangVaoDauSheet(SheetHelper.sheetName.sheetHose);
    sheetHelper.ghiDuLieuVaoDayTheoVung([[ngayHienTai, duLieuNhanVe?.priceClose, tiLeThayDoi / 100, thanhKhoan]], SheetHelper.sheetName.sheetHose, 'A1:D1');
  } else {
    console.log('No action required');
  }
}

function layThongTinCoBan(): void {
  const sheetHelper = new SheetHelper();
  const danhSachMa: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);
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
  const DANH_SACH_MA: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);
  const DANH_SACH_MA_SHEET_GIA = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetGia, 1);
  const MANG_COT_NGAY = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetGia, 'A').slice(1);
  const DATE: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B1');
  // bo cot dau la ngay & so thu tu bat dau tu 0 => +2
  const VI_TRI_NGAY_CAN_TIM_TRONG_COT_NGAY = MANG_COT_NGAY.indexOf(DATE) + 2;
  const MANG_GIA_SHEET_GIA = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetGia, VI_TRI_NGAY_CAN_TIM_TRONG_COT_NGAY);
  let index = 2;
  const datas = DANH_SACH_MA_SHEET_GIA.map((code, i) => ({
    Symbol: code,
    ClosePrice: MANG_GIA_SHEET_GIA[i],
  }));
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
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(url);

    object.data.forEach((element: IResponseVndirect) => {
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
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(url);
    object.data.forEach((element: IResponseVndirect) => {
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
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4';
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(url);
    object.data.forEach((element: IResponseVndirect) => {
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
  const duLieuCotThamChieu = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A').slice(1);

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.kichThuocMangPhu) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetHelper.kichThuocMangPhu).join(',')}`;
    const object = httpHelper.sendGetRequest(url);
    object.data.forEach((element: IResponseVndirect) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = sheetHelper.layViTriCotThamChieu(tenMa, duLieuCotThamChieu, 2);
      sheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.sheetName.sheetDuLieu, `I${vitri}:I${vitri}`);
    });
  }
}

export function duLieuTam(): void {
  const sheetHelper = new SheetHelper();
  sheetHelper.layDuLieuTrongCot('TRUY VAN', 'A').slice(1).forEach((date: string) => {
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
    const url = `https://api-finfo.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${headers.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
    const object = httpHelper.sendGetRequest(url);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: IResponseVndirect) => {
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
    const url = `https://api-finfo.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${headers.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
    const object = httpHelper.sendGetRequest(url);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: IResponseVndirect) => {
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
    const url = `https://api-finfo.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${headers.join(',')}~date:gte:${date}~date:lte:${date}`;
    const object = httpHelper.sendGetRequest(url);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + date]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: IResponseVndirect) => {
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
function layGiaHangNgay(sheetName = SheetHelper.sheetName.sheetGia, ngayMoiNhatSheetHose = new SheetHelper().layDuLieuTrongO(SheetHelper.sheetName.sheetHose, 'A1')) {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const headers: string[] = sheetHelper.layDuLieuTrongHang(SheetHelper.sheetName.sheetKhoiLuong, 1).slice(1);
  const hangCuoi = sheetHelper.laySoHangTrongSheet(sheetName);
  const duLieuNgayMoiNhat = sheetHelper.layDuLieuTrongO(sheetName, 'A' + hangCuoi);
  if (duLieuNgayMoiNhat !== ngayMoiNhatSheetHose) {
    const url = `https://api-finfo.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${headers.join(',')}~date:gte:${ngayMoiNhatSheetHose}~date:lte:${ngayMoiNhatSheetHose}`;
    const object = httpHelper.sendGetRequest(url);
    if (object?.data.length > 0) {
      const header = sheetHelper.layDuLieuTrongHang(sheetName, 1);
      sheetHelper.ghiDuLieuVaoDay([["'" + ngayMoiNhatSheetHose]], sheetName, hangCuoi + 1, 1);
      object.data.map((item: IResponseVndirect) => {
        for (let i = 0; i < header.length; i++) {
          if (item.close !== undefined && item.code === header[i]) {
            sheetHelper.ghiDuLieuVaoDay([[item.close * 1000]], sheetName, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy giá hàng ngày thành công');
  } else {
    console.log('done');
  }
}

/**
 * @customfunction
 */
function LAY_THONG_TIN_DANH_MUC_DC(url: string) {
  const httpHelper = new HttpHelper();
  const sheetHelper = new SheetHelper();
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const result: string[][] = [];
  const response = httpHelper.sendRequest(url);
  const data = response.returnValue.top10Holding;
  const capNhatLuc = DateHelper.doiDinhDangNgayISO(response.returnValue.tradingDate, defaultFormat);
  data.forEach((element: IResponseDC) => {
    const tenMa = element.assetId ?? '_';
    const nhomNganh = element.translation?.vi?.sectorLevel ?? '_';
    const tyLe = element.weight ?? 0;
    const sanGD = element.exchange ?? '_';
    result.push([tenMa, nhomNganh, sanGD, String(tyLe), capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
function LAY_THONG_TIN_TAI_SAN_DC(url: string) {
  const httpHelper = new HttpHelper();
  const sheetHelper = new SheetHelper();
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const result: string[][] = [];
  const response = httpHelper.sendRequest(url);
  const data = response.returnValue.allocationBySectors;
  const capNhatLuc = DateHelper.doiDinhDangNgayISO(response.returnValue.tradingDate, defaultFormat);
  data.forEach((element: IResponseDC) => {
    const tenTaiSan = element.translation?.vi?.industryLevel2 ?? '_';
    const tyLe = element.fundWeight?.VF1 ?? element.fundWeight?.VF4 ?? '_';
    result.push([tenTaiSan, String(tyLe), capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
function LAY_SU_KIEN() {
  const result: string[][] = [];
  const content = UrlFetchApp.fetch(`https://hontrang.github.io/google-sheet/`).getContentText();
  const $ = Cheerio.load(content);
  let date: string;
  $('table#calendar>thead.table-header,table#calendar>tbody').each(function () {
    if ($(this).attr('class') !== undefined) {
      date = $(this).find('tr>th:nth-child(1)').text().trim();
    } else {
      $(this)
        .children('tr')
        .each(function () {
          const timeValue = $(this).find('td:nth-child(1)').text().trim() || '00:00 AM';
          const thoigian = DateHelper.doiDinhDangNgay(`${date} ${timeValue}`, 'EEEE MMMM dd yyyy hh:mm a', 'EEEE yyyy/MM/dd hh:mm a', { locale: 'vi-VN' });
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

function layBaoCaoDC() {
  const httpHelper = new HttpHelper();
  const sheetHelper = new SheetHelper();
  const result: string[][] = [];
  const url = `https://www.dragoncapital.com.vn/individual/vi/webruntime/api/apex/execute?language=vi&asGuest=true&htmlEncode=false`;
  const option = JSON.stringify({
    "namespace": "",
    "classname": "@udd/01pJ2000000CgR7",
    "method": "getDocumentContentsV2",
    "isContinuation": false,
    "params": {
      "siteId": "0DMJ2000000oLukOAE",
      "fundCodeOrReportCode": "VF1",
      "documentType": null,
      "targetYear": DateHelper.layNamHienTaiAsString(),
      "language": "vi"
    },
    "cacheable": false
  });

  const configs = {
    "method": "post",
    "contentType": "application/json",
    "payload": option,
    "headers": {
      'Content-Type': 'application/json',
      'Cookie': 'CookieConsentPolicy=0:1; LSKey-c$CookieConsentPolicy=0:1'
    },
    "maxBodyLength": "Infinity",
  };

  const response = httpHelper.sendPostRequest(url, configs);
  const data = response.returnValue;
  const danhSachBaoCao: IResponseDC[] = data[5].files;
  danhSachBaoCao.forEach((baoCao) => {
    const tenBaoCao = baoCao.activeFileName__c ?? '_';
    const dlc = baoCao.downloadUrl__c ?? '_';
    const capNhatLuc = baoCao.displayDate__c ?? '_';
    result.push([tenBaoCao, dlc, capNhatLuc]);
  });
  sheetHelper.ghiDuLieuVaoDay(result, SheetHelper.sheetName.sheetDC, 2, 17);
}

function layThongTinPhaiSinh() {
  const result: string[][] = [];
  const httpHelper = new HttpHelper();
  const sheetHelper = new SheetHelper();
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const url = `https://api-finfo.vndirect.com.vn/v4/derivatives?q=underlyingType:INDEX~status:LISTED&size=10000`;
  const response = httpHelper.sendGetRequest(url);
  const datas: IResponseVndirect[] = response.data.filter((d: IResponseVndirect) => new Date(d.expiryDate ?? '').getTime() > Date.now())
    .sort((a: IResponseVndirect, b: IResponseVndirect) =>
      new Date(a.expiryDate ?? '').getTime() - new Date(b.expiryDate ?? '').getTime()
    );
  datas.forEach((data) => {
    result.push([DateHelper.doiDinhDangNgay(data.expiryDate ?? '_', defaultFormat, 'EEEE yyyy/MM/dd', { locale: 'vi-VN' })]);
  });
  sheetHelper.ghiDuLieuVaoDay(result, SheetHelper.sheetName.sheetDuLieu, 2, 74);
}