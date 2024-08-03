/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';

function getDataHose(): void {
  const DANH_SACH_MA: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  const mangDuLieuChinh: Array<[string]> = [];
  const url = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
  const response = HttpHelper.sendGetRequest(url);
  for (const element of DANH_SACH_MA) {
    const price = response.filter((object: { sym: string }) => object.sym === element).map((object: { lastPrice: number }) => object.lastPrice * 1000);
    mangDuLieuChinh.push([price]);
  }
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'B');
}

function layTinTucSheetBangThongTin(): void {
  const mangDuLieuChinh: Array<[string, string, string, string, string, string]> = [];
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_CAU_HINH, 'E');
  danhSachMa.forEach((tenMa: string) => {
    const url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content: string = UrlFetchApp.fetch(url).getContentText();
    const $ = Cheerio.load(content);
    $('a').each(function (this: any) {
      const title: string = $(this).attr('title') ?? '';
      const link: string = 'https://s.cafef.vn' + ($(this).attr('href') ?? '');
      const date: string = $(this).siblings('span').text().substring(0, 10);
      mangDuLieuChinh.push([tenMa, title, '', link, '', date]);
    });
  });
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_BANG_THONG_TIN, 35, 'A');
}

function layThongTinChiTietMa(): void {
  SheetHelper.ghiDuLieuVaoO('...', SheetHelper.SHEET_CHI_TIET_MA, 'J2');
  SheetHelper.xoaDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'P', 3, 2);
  const tenMa: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F1');

  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);

  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh();
  layThongTinCoDong(tenMa);
  layThongTongSoLuongCoPhieuDangNiemYet(tenMa);
  layThongTinCoTuc(tenMa);
  layHeSoBetaVaFreeFloat(tenMa);
  ZChartHelper.updateChart();
  LogHelper.logTime(SheetHelper.SHEET_CHI_TIET_MA, 'J2');
  Logger.log('Hàm layThongTinChiTietMa chạy thành công');
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa: string): void {
  const fromDate: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F2');
  const toDate: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'H2');
  const mangDuLieuChinh: Array<[string, number, number]> = [];
  mangDuLieuChinh.push(['chi tiết mã', 0, 0]);
  const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = HttpHelper.sendGetRequest(url);
  const datas = object.data;
  for (const element of datas) {
    mangDuLieuChinh.push([element.date, element.close * 1000, element.nmVolume]);
  }
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 1, 'P');
}

function layTinTucSheetChiTietMa(tenMa: string): void {
  const baseUrl = 'https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx';
  const mangDuLieuChinh: Array<[string, string, string, string]> = [];
  const queryUrl = `${baseUrl}?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  const $ = Cheerio.load(content);
  $('a').each(function (this: any) {
    const title = $(this).attr('title') ?? '';
    const link = `${baseUrl}${$(this).attr('href') ?? ''}`;
    const date = $(this).siblings('span').text().substring(0, 10);
    mangDuLieuChinh.push([tenMa.toUpperCase(), title, link, date]);
  });
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'AH');
}

function layBaoCaoTaiChinh(): void {
  const mangDuLieuChinh: Array<[string, string, string, string]> = [];
  const tenMa: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F1');
  const queryUrl = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  const $ = Cheerio.load(content);
  $('#divDocument>div>table>tbody>tr').each(function (this: any) {
    const title: string = $(this).children('td:nth-child(1)').text();
    const date: string = $(this).children('td:nth-child(2)').text();
    const link: string = $(this).children('td:nth-child(3)').children('a').attr('href') ?? '';
    mangDuLieuChinh.push([tenMa.toUpperCase(), title, date, link]);
  });
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh.slice(1, 11), SheetHelper.SHEET_DU_LIEU, 18, 'AH');
}

function layBaoCaoPhanTich(tenMa: string): void {
  const mangDuLieuChinh: [string, string, string, string, string][] = [];
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = HttpHelper.sendPostRequest(url);

  const datas = object.Data;
  for (const element of datas) {
    const sourceName: string = element.SourceName ?? '____';
    const title: string = element.Title ?? '____';
    const reportTypeName: string = element.ReportTypeName ?? '____';
    const lastUpdate: string = element.LastUpdate ?? '____';
    const url: string = element.Url ?? '____';
    mangDuLieuChinh.push([sourceName, title, reportTypeName, lastUpdate, url]);
  }

  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'AL');
}

function layThongTinCoDong(tenMa: string): void {
  const url = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;
  const object = HttpHelper.sendRequest(url);
  const mangDuLieuChinh: Array<[string, string, string]> = object.listShareHolder.map(({ ticker, name, ownPercent }: { ticker: string; name: string; ownPercent: string }) => [
    ticker,
    name,
    ownPercent
  ]);

  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'AD');
}

function layThongTinCoTuc(tenMa: string): void {
  const mangDuLieuChinh: [string, string, string][] = [];

  const OPTIONS_CO_TUC: URLFetchRequestOptions = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    payload: JSON.stringify({
      tickers: [`${tenMa}`],
      page: 0,
      size: 10
    })
  };
  const url = `https://api.simplize.vn/api/company/separate-share/list-tickers`;
  const response = HttpHelper.sendRequest(url, OPTIONS_CO_TUC);

  const datas = response.data;
  for (const element of datas) {
    const content: string = element.content ?? '____';
    const date: string = element.date ?? '____';
    mangDuLieuChinh.push([content, '', date]);
  }

  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 18, 'AO');
}

function layThongTongSoLuongCoPhieuDangNiemYet(tenMa: string): void {
  const market = 'HOSE';
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=1&lookupRequest.pageSize=1000&lookupRequest.symbol=${tenMa}`;
  const token = HttpHelper.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: token,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    }
  };
  const object = HttpHelper.sendRequest(url, OPTION);
  const mangDuLieuChinh: Array<[string]> = [];
  mangDuLieuChinh.push([object.data[0].RepeatedInfo[0].ListedShare]);
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_CHI_TIET_MA, 18, 'H');
}

function layHeSoBetaVaFreeFloat(tenMa = 'FRT') {
  const fromDate: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F2');
  const URL = `https://finfo-api.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
  const response = HttpHelper.sendGetRequest(URL);
  const datas = response.data;
  for (const element of datas) {
    if (element.ratioCode === 'BETA') {
      const value = element.value;
      SheetHelper.ghiDuLieuVaoO(value, SheetHelper.SHEET_CHI_TIET_MA, 'E18');
    }
    if (element.ratioCode === 'FREEFLOAT') {
      const value = element.value;
      SheetHelper.ghiDuLieuVaoO(value, SheetHelper.SHEET_CHI_TIET_MA, 'J16');
    }
  }
}

function batSukienSuaThongTinO(e: any) {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (e.range.getA1Notation() === 'F1' && sheet.getName() === SheetHelper.SHEET_CHI_TIET_MA) {
    layThongTinChiTietMa();
  }
}
