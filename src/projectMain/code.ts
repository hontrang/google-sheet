/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

import { SheetHelper, SHEET_NAME } from '../utility/SheetHelper';
import { HttpHelper } from '../utility/HttpHelper';
import { DateHelper } from '../utility/DateHelper';
import { LogHelper } from '../utility/LogHelper';

function getDataHose(): void {
  const DANH_SACH_MA: string[] = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'A');
  let indexSheetDuLieu = 2;
  let indexSheetThamChieu = 4;
  const url = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
  const response = new HttpHelper().sendGetRequest(url);

  const priceMap: Record<string, number> = {};
  response.forEach((object: { sym: string, lastPrice: number }) => {
    priceMap[object.sym] = object.lastPrice * 1000;
  });

  new SheetHelper().xoaDuLieuTrongCot(SHEET_NAME.sheetThamChieu, 'A', 1, 4);

  for (const element of DANH_SACH_MA) {
    const price = priceMap[element] || 0;
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[element]], SHEET_NAME.sheetThamChieu, `A${indexSheetThamChieu}`);
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[price]], SHEET_NAME.sheetDuLieu, `B${indexSheetDuLieu}`);
    indexSheetDuLieu++;
    indexSheetThamChieu++;
  }
}

function layTinTucSheetBangThongTin(): void {
  const mangDuLieuChinh: Array<[string, string, string, string, string, string]> = [];
  const danhSachMa: string[] = new SheetHelper().layDuLieuTrongCot(SHEET_NAME.sheetCauHinh, 'E');
  const baseUrl = 'https://s.cafef.vn';
  danhSachMa.forEach((tenMa: string) => {
    const url = `${baseUrl}/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content: string = UrlFetchApp.fetch(url).getContentText();
    const $ = Cheerio.load(content);
    $('a').each(function (this: any) {
      const title: string = $(this).attr('title') ?? '';
      const link: string = baseUrl + ($(this).attr('href') ?? '');
      const date: string = $(this).siblings('span').text().substring(0, 10);
      mangDuLieuChinh.push([tenMa, title, '', link, '', date]);
    });
  });
  new SheetHelper().ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SHEET_NAME.sheetBangThongTin, 35, 'A');
}

function layThongTinChiTietMa(): void {
  new SheetHelper().ghiDuLieuVaoO('...', SHEET_NAME.sheetChiTietMa, 'J2');
  new SheetHelper().xoaDuLieuTrongCot(SHEET_NAME.sheetDuLieu, 'P', 3, 2);
  const tenMa: string = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetChiTietMa, 'F1');

  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);

  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh(tenMa);
  layThongTinCoDong(tenMa);
  layThongTongSoLuongCoPhieuDangNiemYet(tenMa);
  layThongTinCoTuc(tenMa);
  layHeSoBetaVaFreeFloat(tenMa);
  LogHelper.logTime(SHEET_NAME.sheetChiTietMa, 'J2');
  Logger.log('Hàm layThongTinChiTietMa chạy thành công');
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa = 'FRT'): void {
  const fromDate: string = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetChiTietMa, 'F2');
  const toDate: string = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetChiTietMa, 'H2');
  let index = 2;
  new SheetHelper().ghiDuLieuVaoDayTheoVung([['chi tiết mã', '', '']], SHEET_NAME.sheetDuLieu, 'P1:R1');
  const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = new HttpHelper().sendGetRequest(url);
  const datas = object.data;
  for (const element of datas) {
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[element.date, element.close * 1000, element.nmVolume]], SHEET_NAME.sheetDuLieu, `P${index}:R${index}`);
    index++;
  }
}

function layTinTucSheetChiTietMa(tenMa = 'FRT'): void {
  const baseUrl = 'https://s.cafef.vn';
  const queryUrl = `${baseUrl}/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  const DEFAULT_FORMAT = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B6');
  let index = 2;
  const $ = Cheerio.load(content);
  $('a').each(function (this: any) {
    const title = $(this).attr('title') ?? '';
    const link = `${baseUrl}${$(this).attr('href') ?? ''}`;
    const date = DateHelper.doiDinhDangNgay($(this).siblings('span').text().substring(0, 10), 'DD/MM/YYYY', DEFAULT_FORMAT);
    new SheetHelper().ghiDuLieuVaoDayTheoVung([[tenMa.toUpperCase(), title, link, date]], SHEET_NAME.sheetDuLieu, `AH${index}:AK${index}`);
    index++;
  });
}

// lấy 10 báo cáo tài chính đầu tiên
function layBaoCaoTaiChinh(tenMa = 'FRT'): void {
  const queryUrl = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  let index = 18;
  const $ = Cheerio.load(content);
  $('#divDocument>div>table>tbody>tr').each(function (this: any) {
    const title: string = $(this).children('td:nth-child(1)').text();
    const date: string = $(this).children('td:nth-child(2)').text();
    const link: string = $(this).children('td:nth-child(3)').children('a').attr('href') ?? '';
    if (title !== 'Loại báo cáo' && index < 28) {
      new SheetHelper().ghiDuLieuVaoDayTheoVung([[tenMa.toUpperCase(), title, date, link]], SHEET_NAME.sheetDuLieu, `AH${index}:AK${index}`);
      index++;
    }
  });
}

function layBaoCaoPhanTich(tenMa = 'FRT'): void {
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = new HttpHelper().sendPostRequest(url);
  const DEFAULT_FORMAT = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B6');
  let index = 2;
  const datas = object.Data;
  for (const element of datas) {
    const sourceName: string = element.SourceName ?? '_';
    const title: string = element.Title ?? '_';
    const reportTypeName: string = element.ReportTypeName ?? '_';
    const lastUpdate: string = element.LastUpdate ?? '_';
    const url: string = element.Url ?? '_';
    new SheetHelper().ghiDuLieuVaoDayTheoVung(
      [[sourceName, title, reportTypeName, DateHelper.doiDinhDangNgay(lastUpdate, 'DD/MM/YYYY', DEFAULT_FORMAT), url]],
      SHEET_NAME.sheetDuLieu,
      `AL${index}:AP${index}`
    );
    index++;
  }
}

function layThongTinCoDong(tenMa = 'FRT'): void {
  const url = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
  const object = new HttpHelper().sendRequest(url);
  const mangDuLieuChinh = object.listShareHolder.map(({ ticker, name, ownPercent }: { ticker: string; name: string; ownPercent: string }) => [ticker, name, ownPercent]);

  new SheetHelper().ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SHEET_NAME.sheetDuLieu, 2, 'AD');
}

function layThongTinCoTuc(tenMa = 'FRT'): void {
  const DEFAULT_FORMAT = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetCauHinh, 'B6');
  let index = 18;
  const OPTIONS_CO_TUC: URLFetchRequestOptions = {
    method: 'post',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    payload: JSON.stringify({
      tickers: [`${tenMa}`],
      page: 0,
      size: 10
    })
  };
  const url = `https://api.simplize.vn/api/company/separate-share/list-tickers`;
  const response = new HttpHelper().sendRequest(url, OPTIONS_CO_TUC);

  const datas = response.data;
  for (const element of datas) {
    const content: string = element.content ?? '____';
    const date: string = element.date ?? '____';
    new SheetHelper().ghiDuLieuVaoDayTheoVung(
      [[content, '', DateHelper.doiDinhDangNgay(date, 'DD/MM/YYYY', DEFAULT_FORMAT)]],
      SHEET_NAME.sheetDuLieu,
      `AO${index}:AQ${index}`
    );
    index++;
  }
}

function layThongTongSoLuongCoPhieuDangNiemYet(tenMa = 'FRT'): void {
  const market = 'HOSE';
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=1&lookupRequest.pageSize=1000&lookupRequest.symbol=${tenMa}`;
  const token = new HttpHelper().getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = new HttpHelper().sendRequest(url, OPTION);
  new SheetHelper().ghiDuLieuVaoDayTheoTen([[[object.data[0].RepeatedInfo[0].ListedShare]]], SHEET_NAME.sheetChiTietMa, 18, 'H');
}

function layHeSoBetaVaFreeFloat(tenMa = 'FRT') {
  const fromDate: string = new SheetHelper().layDuLieuTrongO(SHEET_NAME.sheetChiTietMa, 'F2');
  const URL = `https://finfo-api.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
  const response = new HttpHelper().sendGetRequest(URL);
  const datas = response.data;
  for (const element of datas) {
    if (element.ratioCode === 'BETA') {
      const value = element.value;
      new SheetHelper().ghiDuLieuVaoO(value, SHEET_NAME.sheetChiTietMa, 'E18');
    }
    if (element.ratioCode === 'FREEFLOAT') {
      const value = element.value;
      new SheetHelper().ghiDuLieuVaoO(value, SHEET_NAME.sheetChiTietMa, 'J16');
    }
  }
}

function batSukienSuaThongTinO(e: any) {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (e.range.getA1Notation() === 'F1' && sheet.getName() === SHEET_NAME.sheetChiTietMa) {
    layThongTinChiTietMa();
  }
}
