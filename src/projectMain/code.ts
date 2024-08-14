/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

import { HttpHelper } from '../utility/HttpHelper';
import { DateHelper } from '../utility/DateHelper';
import { LogHelper } from '../utility/LogHelper';
import { SheetHelper } from '@utils/SheetHelper';

function getDataHose(): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const DANH_SACH_MA: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');
  let indexSheetDuLieu = 2;
  let indexSheetThamChieu = 4;
  const url = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
  const response = httpHelper.sendGetRequest(url);

  const priceMap: Record<string, number> = {};
  response.forEach((object: { sym: string, lastPrice: number }) => {
    priceMap[object.sym] = object.lastPrice * 1000;
  });

  sheetHelper.xoaDuLieuTrongCot(SheetHelper.sheetName.sheetThamChieu, 'A', 1, 4);

  for (const element of DANH_SACH_MA) {
    const price = priceMap[element] || 0;
    sheetHelper.ghiDuLieuVaoDayTheoVung([[element]], SheetHelper.sheetName.sheetThamChieu, `A${indexSheetThamChieu}`);
    sheetHelper.ghiDuLieuVaoDayTheoVung([[price]], SheetHelper.sheetName.sheetDuLieu, `B${indexSheetDuLieu}`);
    indexSheetDuLieu++;
    indexSheetThamChieu++;
  }
}

function layTinTucSheetBangThongTin(): void {
  const sheetHelper = new SheetHelper();
  const mangDuLieuChinh: Array<[string, string, string, string, string, string]> = [];
  const danhSachMa: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetCauHinh, 'E');
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
  sheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.sheetName.sheetBangThongTin, 35, 'A');
}

function layThongTinChiTietMa(): void {
  const sheetHelper = new SheetHelper();
  sheetHelper.ghiDuLieuVaoO('...', SheetHelper.sheetName.sheetChiTietMa, 'J2');
  sheetHelper.xoaDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'P', 3, 2);
  const tenMa: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'F1');

  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);

  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh(tenMa);
  layThongTinCoDong(tenMa);
  layThongTongSoLuongCoPhieuDangNiemYet(tenMa);
  layThongTinCoTuc(tenMa);
  layHeSoBetaVaFreeFloat(tenMa);
  LogHelper.logTime(SheetHelper.sheetName.sheetChiTietMa, 'J2');
  Logger.log('Hàm layThongTinChiTietMa chạy thành công');
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const fromDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'F2');
  const toDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'H2');
  let index = 2;
  sheetHelper.ghiDuLieuVaoDayTheoVung([['chi tiết mã', '', '']], SheetHelper.sheetName.sheetDuLieu, 'P1:R1');
  const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = httpHelper.sendGetRequest(url);
  const datas = object.data;
  for (const element of datas) {
    sheetHelper.ghiDuLieuVaoDayTheoVung([[element.date, element.close * 1000, element.nmVolume]], SheetHelper.sheetName.sheetDuLieu, `P${index}:R${index}`);
    index++;
  }
}

function layTinTucSheetChiTietMa(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const baseUrl = 'https://s.cafef.vn';
  const queryUrl = `${baseUrl}/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  const DEFAULT_FORMAT = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  let index = 2;
  const $ = Cheerio.load(content);
  $('a').each(function (this: any) {
    const title = $(this).attr('title') ?? '';
    const link = `${baseUrl}${$(this).attr('href') ?? ''}`;
    const date = DateHelper.doiDinhDangNgay($(this).siblings('span').text().substring(0, 10), 'DD/MM/YYYY', DEFAULT_FORMAT);
    sheetHelper.ghiDuLieuVaoDayTheoVung([[tenMa.toUpperCase(), title, link, date]], SheetHelper.sheetName.sheetDuLieu, `AH${index}:AK${index}`);
    index++;
  });
}

// lấy 10 báo cáo tài chính đầu tiên
function layBaoCaoTaiChinh(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const queryUrl = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  let index = 18;
  const $ = Cheerio.load(content);
  $('#divDocument>div>table>tbody>tr').each(function (this: any) {
    const title: string = $(this).children('td:nth-child(1)').text();
    const date: string = $(this).children('td:nth-child(2)').text();
    const link: string = $(this).children('td:nth-child(3)').children('a').attr('href') ?? '';
    if (title !== 'Loại báo cáo' && index < 28) {
      sheetHelper.ghiDuLieuVaoDayTheoVung([[tenMa.toUpperCase(), title, date, link]], SheetHelper.sheetName.sheetDuLieu, `AH${index}:AK${index}`);
      index++;
    }
  });
}

function layBaoCaoPhanTich(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = httpHelper.sendPostRequest(url);
  const DEFAULT_FORMAT = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  let index = 2;
  const datas = object.Data;
  for (const element of datas) {
    const sourceName: string = element.SourceName ?? '_';
    const title: string = element.Title ?? '_';
    const reportTypeName: string = element.ReportTypeName ?? '_';
    const lastUpdate: string = element.LastUpdate ?? '_';
    const url: string = element.Url ?? '_';
    sheetHelper.ghiDuLieuVaoDayTheoVung(
      [[sourceName, title, reportTypeName, DateHelper.doiDinhDangNgay(lastUpdate, 'DD/MM/YYYY', DEFAULT_FORMAT), url]],
      SheetHelper.sheetName.sheetDuLieu,
      `AL${index}:AP${index}`
    );
    index++;
  }
}

function layThongTinCoDong(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const url = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
  const object = httpHelper.sendRequest(url);
  const mangDuLieuChinh = object.listShareHolder.map(({ ticker, name, ownPercent }: { ticker: string; name: string; ownPercent: string }) => [ticker, name, ownPercent]);

  sheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.sheetName.sheetDuLieu, 2, 'AD');
}

function layThongTinCoTuc(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const DEFAULT_FORMAT = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
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
  const response = httpHelper.sendRequest(url, OPTIONS_CO_TUC);

  const datas = response.data;
  for (const element of datas) {
    const content: string = element.content ?? '____';
    const date: string = element.date ?? '____';
    sheetHelper.ghiDuLieuVaoDayTheoVung(
      [[content, '', DateHelper.doiDinhDangNgay(date, 'DD/MM/YYYY', DEFAULT_FORMAT)]],
      SheetHelper.sheetName.sheetDuLieu,
      `AO${index}:AQ${index}`
    );
    index++;
  }
}

function layThongTongSoLuongCoPhieuDangNiemYet(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const market = 'HOSE';
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=1&lookupRequest.pageSize=1000&lookupRequest.symbol=${tenMa}`;
  const token = httpHelper.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = httpHelper.sendRequest(url, OPTION);
  sheetHelper.ghiDuLieuVaoDayTheoTen([[[object.data[0].RepeatedInfo[0].ListedShare]]], SheetHelper.sheetName.sheetChiTietMa, 18, 'H');
}

function layHeSoBetaVaFreeFloat(tenMa = 'FRT') {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const fromDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'F2');
  const URL = `https://finfo-api.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
  const response = httpHelper.sendGetRequest(URL);
  const datas = response.data;
  for (const element of datas) {
    if (element.ratioCode === 'BETA') {
      const value = element.value;
      sheetHelper.ghiDuLieuVaoO(value, SheetHelper.sheetName.sheetChiTietMa, 'E18');
    }
    if (element.ratioCode === 'FREEFLOAT') {
      const value = element.value;
      sheetHelper.ghiDuLieuVaoO(value, SheetHelper.sheetName.sheetChiTietMa, 'J16');
    }
  }
}

function batSukienSuaThongTinO(e: any) {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (e.range.getA1Notation() === 'F1' && sheet.getName() === SheetHelper.sheetName.sheetChiTietMa) {
    layThongTinChiTietMa();
  }
}
