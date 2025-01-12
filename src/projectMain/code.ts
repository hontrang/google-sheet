/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

import { HttpHelper } from '@utils/HttpHelper';
import { DateHelper } from '@utils/DateHelper';
import { LogHelper } from '@utils/LogHelper';
import { SheetHelper } from '@utils/SheetHelper';
import { ZchartHelper } from '@utils/zChartUtil';
import { ResponseSimplize, ResponseVietStock, ResponseVndirect } from '@src/types/types';

function getDataHose(): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const DANH_SACH_MA: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'A');
  let indexSheetDuLieu = 2;
  let indexSheetThamChieu = 4;
  const url = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(',')}`;
  const response = httpHelper.sendGetRequest(url);

  const priceMap: Record<string, number> = {};
  response.forEach((object: { sym: string; lastPrice: number }) => {
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
  LogHelper.logTime(SheetHelper.sheetName.sheetBangThongTin, 'I1');
}

function layTinTucSheetBangThongTin(): void {
  const sheetHelper = new SheetHelper();
  const mangDuLieuChinh: Array<[string, string, string, string, string, string]> = [];
  const danhSachMa: string[] = sheetHelper.layDuLieuTrongCot(SheetHelper.sheetName.sheetCauHinh, 'E');
  const baseUrl = 'https://cafef.vn';
  danhSachMa.forEach((tenMa: string) => {
    const url = `${baseUrl}/du-lieu/tin-doanh-nghiep/${tenMa}/event.chn`;
    const content: string = UrlFetchApp.fetch(url).getContentText();
    const $ = Cheerio.load(content);
    $('a.docnhanhTitle').each(function (i) {
      if (i < 10) {
        const title: string = $(this).attr('title') ?? '_';
        const link: string = baseUrl + ($(this).attr('href') ?? '_');
        const date: string = $(this).siblings('span').text().substring(0, 10);
        mangDuLieuChinh.push([tenMa, title, '', link, '', date]);
      }
    });
  });
  sheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.sheetName.sheetBangThongTin, 35, 'A');
}

function layThongTinChiTietMa(): void {
  const sheetHelper = new SheetHelper();
  sheetHelper.ghiDuLieuVaoO('...', SheetHelper.sheetName.sheetChiTietMa, 'I1');
  sheetHelper.xoaDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, 'P', 3, 2);
  const tenMa: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'B1');

  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);

  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh(tenMa);
  layThongTinCoDong(tenMa);
  layTongSoLuongCoPhieuDangNiemYet(tenMa);
  layThongTinCoTuc(tenMa);
  layHeSoBetaVaFreeFloat(tenMa);
  layDonViKiemToan(tenMa);
  layChiTietBaoCaoTaiChinh(tenMa);
  layThongTinTraiPhieu(tenMa);
  ZchartHelper.updateChart();
  LogHelper.logTime(SheetHelper.sheetName.sheetChiTietMa, 'I1');
  Logger.log('Hàm layThongTinChiTietMa chạy thành công');
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const fromDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B17');
  const toDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B18');
  let index = 2;
  sheetHelper.ghiDuLieuVaoDayTheoVung([['chi tiết mã', '', '']], SheetHelper.sheetName.sheetDuLieu, 'P1:R1');
  const url = `https://api-finfo.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = httpHelper.sendGetRequest(url);
  const datas: [ResponseVndirect] = object.data;
  for (const element of datas) {
    const ngay = element.date ?? '_';
    const gia = element.close ?? 0;
    const khoiLuong = element.nmVolume ?? '_';
    sheetHelper.ghiDuLieuVaoDayTheoVung([[ngay, gia * 1000, khoiLuong]], SheetHelper.sheetName.sheetDuLieu, `P${index}:R${index}`);
    index++;
  }
}

function layTinTucSheetChiTietMa(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const baseUrl = 'https://s.cafef.vn';
  const queryUrl = `${baseUrl}/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content: string = UrlFetchApp.fetch(queryUrl).getContentText();
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  let index = 2;
  const $ = Cheerio.load(content);
  $('a').each(function () {
    const title = $(this).attr('title') ?? '';
    const link = `${baseUrl}${$(this).attr('href') ?? ''}`;
    const date = DateHelper.doiDinhDangNgay($(this).siblings('span').text().substring(0, 10), 'dd/MM/yyyy', defaultFormat);
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
  $('#divDocument>div>table>tbody>tr').each(function () {
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
  const url = `https://api.simplize.vn/api/company/analysis-report/list?ticker=${tenMa}&isWl=false&page=0&size=10`;
  const object = httpHelper.sendGetRequest(url);
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  let index = 2;
  const datas: [ResponseSimplize] = object.data;
  for (const element of datas) {
    const sourceName: string = element.source ?? '_';
    const title: string = element.title ?? '_';
    const targetPrice: number = element.targetPrice ?? 0;
    const lastUpdate: string = element.issueDate ?? '_';
    const url: string = element.attachedLink ?? '_';
    const khuyenNghi: string = element.recommend ?? 'Không Có';
    sheetHelper.ghiDuLieuVaoDayTheoVung(
      [[sourceName, title, `${targetPrice}`, DateHelper.doiDinhDangNgay(lastUpdate, 'dd/MM/yyyy', defaultFormat), url, khuyenNghi]],
      SheetHelper.sheetName.sheetDuLieu,
      `AL${index}:AQ${index}`
    );
    index++;
  }
}

function layThongTinCoDong(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const url = `https://api.simplize.vn/api/company/ownership/shareholder-fund-details/${tenMa}`;
  const response = httpHelper.sendRequest(url);
  const danhSach = response.data.shareholderDetails.slice(0, 10);
  const mangDuLieuChinh = danhSach.map(({ investorFullName, pctOfSharesOutHeld, changeValue, countryOfInvestor }: ResponseSimplize) => [investorFullName, `${pctOfSharesOutHeld}`, `${changeValue}`, countryOfInvestor]);

  sheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.sheetName.sheetDuLieu, 2, 'AD');
}

function layThongTinCoTuc(tenMa = 'FRT'): void {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
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

  const datas: [ResponseSimplize] = response.data;
  for (const element of datas) {
    const content: string = element.content ?? '_';
    const date: string = element.date ?? '_';
    sheetHelper.ghiDuLieuVaoDayTheoVung([[content, '', DateHelper.doiDinhDangNgay(date, 'dd/MM/yyyy', defaultFormat)]], SheetHelper.sheetName.sheetDuLieu, `AO${index}:AQ${index}`);
    index++;
  }
}

async function layTongSoLuongCoPhieuDangNiemYet(tenMa = 'FRT'): Promise<void> {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const market = 'HOSE';
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=1&lookupRequest.pageSize=1000&lookupRequest.symbol=${tenMa}`;
  const token = await httpHelper.getToken();
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
  const fromDate: string = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B17');
  const URL = `https://api-finfo.vndirect.com.vn/v4/ratios/latest?filter=ratioCode:MARKETCAP,NMVOLUME_AVG_CR_10D,PRICE_HIGHEST_CR_52W,PRICE_LOWEST_CR_52W,OUTSTANDING_SHARES,FREEFLOAT,BETA,PRICE_TO_EARNINGS,PRICE_TO_BOOK,DIVIDEND_YIELD,BVPS_CR,&where=code:${tenMa}~reportDate:gt:${fromDate}&order=reportDate&fields=ratioCode,value`;
  const response = httpHelper.sendGetRequest(URL);
  const datas: [ResponseVndirect] = response.data;
  for (const element of datas) {
    if (element.ratioCode === 'BETA') {
      const value = element.value;
      sheetHelper.ghiDuLieuVaoO(value, SheetHelper.sheetName.sheetChiTietMa, 'E18');
    }
    if (element.ratioCode === 'FREEFLOAT') {
      const value = element.value;
      sheetHelper.ghiDuLieuVaoO(value, SheetHelper.sheetName.sheetChiTietMa, 'E20');
    }
  }
}

function layDonViKiemToan(tenMa = 'FRT') {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const URL = `https://api-finfo.vndirect.com.vn/v4/company_relations?q=code:${tenMa}~relationType:AUDITOR&size=100&sort=year:DESC`;
  const response = httpHelper.sendGetRequest(URL);
  const datas = response.data;
  let index = 29;
  datas.forEach(function (element: ResponseVndirect) {
    sheetHelper.ghiDuLieuVaoDayTheoVung([[element.relationNameVn, '', '', element.year]], SheetHelper.sheetName.sheetDuLieu, `AH${index}:AK${index}`);
    index++;
  });
}
function layChiTietBaoCaoTaiChinh(tenMa = 'FRT') {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const URL = `https://api-finfo.vndirect.com.vn/v4/financial_statements?q=code:${tenMa}~reportType:QUARTER~modelType:1,89,3,91~fiscalDate:${DateHelper.layKiTaiChinhTheoQuy()}&sort=fiscalDate&size=2000`;

  const response = httpHelper.sendGetRequest(URL);
  const datas = response.data;
  if (datas.length <= 0) {
    sheetHelper.ghiDuLieuVaoO('Lỗi dữ liệu', SheetHelper.sheetName.sheetDuLieu, 'AH29');
  } else {
    let index = 29;
    datas.forEach(function (element: ResponseVndirect) {
      // Tiền và tương đương tiền
      if (element.itemCode === 37000) {
        sheetHelper.ghiDuLieuVaoDayTheoVung([[`${element.numericValue}`, '', element.fiscalDate]], SheetHelper.sheetName.sheetDuLieu, `AL${index}:AN${index}`);
        index++;
      }
    });
    index = 32;
    datas.forEach(function (element: ResponseVndirect) {
      // Tổng tài sản
      if (element.itemCode === 12700) {
        sheetHelper.ghiDuLieuVaoDayTheoVung([[`${element.numericValue}`, '', element.fiscalDate]], SheetHelper.sheetName.sheetDuLieu, `AL${index}:AN${index}`);
        index++;
      }
    });
    index = 35;
    datas.forEach(function (element: ResponseVndirect) {
      // Nợ ngắn hạn
      if (element.itemCode === 13100) {
        sheetHelper.ghiDuLieuVaoDayTheoVung([[`${element.numericValue}`, '', element.fiscalDate]], SheetHelper.sheetName.sheetDuLieu, `AL${index}:AN${index}`);
        index++;
      }
    });
  }
}

function vietstock() {
  const sheetHelper = new SheetHelper();
  const tenMa = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'B1');
  const response = UrlFetchApp.fetch(`https://finance.vietstock.vn/${tenMa}/trai-phieu-lien-quan.htm`);
  const htmlContent = response.getContentText();
  const regex = `<input.*name=__RequestVerificationToken.*value=([^ >]+)`;
  const match = RegExp(regex).exec(htmlContent);
  const token = match?.[1] ?? "Lỗi dữ liệu";
  sheetHelper.ghiDuLieuVaoO(token, SheetHelper.sheetName.sheetCauHinh, 'B19');
}

function derivative() {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const url = 'https://finance.vietstock.vn/Data/GetBondRelated';
  const tenMa = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'B1');
  const headers = {
    'Accept': '*/*',
    'Accept-Language': 'vi',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Cookie': 'language=vi-VN; __RequestVerificationToken=HOrcaOV9AyD405rXuf5nsBOVwtGAq27usAhYlUknoiTKB9BeVyBMRxdfbnJSEF-fqekKEKIYAJHq6rPRoQz5H0Sz90s7OdoD7WSNLgfcnUc1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
  };

  const payload = {
    '__RequestVerificationToken': 'CXImyVGvqSldrS6OJYhhxioGCJYHWNA3Z5DaNLDrwRFZMEKaSTeJWi21Utfue-GpXm3Yb4poiDsRmvrjs01y1rCtZCmd3mfbx7WgjnIMB5Q1',
    'code': `${tenMa}`,
    'orderBy': 'ReleaseDate',
    'orderDir': 'DESC',
    'page': 1,
    'pageSize': 20
  };

  const options = {
    'method': 'post',
    'headers': headers,
    'payload': payload,
    'muteHttpExceptions': true
  };

  const response = httpHelper.sendPostRequest(url, options);
  sheetHelper.ghiDuLieuVaoO(response, SheetHelper.sheetName.sheetCauHinh, 'B20');
}

function layThongTinTraiPhieu(tenMa = 'FRT') {
  const sheetHelper = new SheetHelper();
  const httpHelper = new HttpHelper();
  const result: string[][] = [];
  const defaultFormat = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B6');
  const url = 'https://finance.vietstock.vn/Data/GetBondRelated';
  const headers = {
    'Accept': '*/*',
    'Accept-Language': 'vi',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Cookie': 'language=vi-VN; __RequestVerificationToken=HOrcaOV9AyD405rXuf5nsBOVwtGAq27usAhYlUknoiTKB9BeVyBMRxdfbnJSEF-fqekKEKIYAJHq6rPRoQz5H0Sz90s7OdoD7WSNLgfcnUc1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'
  };

  const payload = {
    '__RequestVerificationToken': 'CXImyVGvqSldrS6OJYhhxioGCJYHWNA3Z5DaNLDrwRFZMEKaSTeJWi21Utfue-GpXm3Yb4poiDsRmvrjs01y1rCtZCmd3mfbx7WgjnIMB5Q1',
    'code': `${tenMa}`,
    'orderBy': 'ReleaseDate',
    'orderDir': 'DESC',
    'page': 1,
    'pageSize': 20
  };

  const options = {
    'method': 'post',
    'headers': headers,
    'payload': payload,
    'muteHttpExceptions': true
  };

  const response = httpHelper.sendPostRequest(url, options);
  sheetHelper.xoaDuLieuTrongCot(SheetHelper.sheetName.sheetDuLieu, `S`, 5, 40, 30);
  response.forEach(function (element: ResponseVietStock) {
    const tenTP = element.KeyCode ?? '_';
    const ngayPhatHanh = DateHelper.doiTuMillisSangNgay(Number(element.ReleaseDate?.replace('/Date(', '').replace(')/', '')), defaultFormat) ?? '_';
    const ngayDenHan = DateHelper.doiTuMillisSangNgay(Number(element.DueDate?.replace('/Date(', '').replace(')/', '')), defaultFormat) ?? '_';
    const menhGia = element.FaceValue ?? 0;
    const khoiLuong = element.IssuaVolume ?? 0;
    result.push([tenTP, ngayPhatHanh, ngayDenHan, `${menhGia}`, `${khoiLuong}`]);
  });
  if (result.length === 0) {
    console.log(`Không có dữ liệu`);
    return;
  };
  // ghi dữ liệu từ ô S40
  sheetHelper.ghiDuLieuVaoDay(result, SheetHelper.sheetName.sheetDuLieu, 40, 19);

}

function batSukienSuaThongTinO(e: GoogleAppsScript.Events.SheetsOnEdit) {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (e.range.getA1Notation() === 'B1' && sheet.getName() === SheetHelper.sheetName.sheetChiTietMa) {
    layThongTinChiTietMa();
  }
}
