import * as Cheerio from 'cheerio';

function getDataHose(): void {
  const DANH_SACH_MA: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
  const mangDuLieuChinh: Array<[string]> = [];
  const URL: string = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(",")}`;
  const response = SheetHttp.sendGetRequest(URL);
  for (const element of DANH_SACH_MA) {
    const price = response.filter((object: { sym: string; }) => object.sym === element).map((object: { lastPrice: number; }) => object.lastPrice * 1000);
    mangDuLieuChinh.push([price]);
  }
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "B");
}

function layTinTucSheetBangThongTin(): void {
  const mangDuLieuChinh: Array<[string, string, string, string, string, string]> = [];
  const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_CAU_HINH, "E");
  danhSachMa.forEach((tenMa: string) => {
    const url: string = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content: string = UrlFetchApp.fetch(url).getContentText();
    const $ = Cheerio.load(content);
    $("a").each(function (this: any) {
      const title: string = $(this).attr("title") ?? "";
      const link: string = "https://s.cafef.vn" + ($(this).attr("href") ?? "");
      const date: string = $(this).siblings("span").text().substring(0, 10);
      mangDuLieuChinh.push([tenMa, title, "", link, "", date]);
    });
  });
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_BANG_THONG_TIN, 31, "A");
}

function layThongTinChiTietMa(): void {
  SheetUtil.ghiDuLieuVaoO("...", SheetUtil.SHEET_CHI_TIET_MA, "J2");
  SheetUtil.xoaDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "P", 3, 2);
  const tenMa: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "F1");

  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);

  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh();
  layThongTinCoDong(tenMa);
  layThongTongSoLuongCoPhieuDangNiemYet(tenMa);
  ZChartUtil.updateChart();
  SheetLog.logTime(SheetUtil.SHEET_CHI_TIET_MA, "J2");
  Logger.log("Hàm layThongTinChiTietMa chạy thành công");
}


function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa: string): void {
  const fromDate: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "F2");
  const toDate: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "H2");
  const mangDuLieuChinh: Array<[string, number, number]> = [];
  mangDuLieuChinh.push(["chi tiết mã", 0, 0])
  const URL: string = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = SheetHttp.sendGetRequest(URL);
  const datas = object.data;
  for (const element of datas) {
    mangDuLieuChinh.push([element.date, element.close * 1000, element.nmVolume]);
  }
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 1, "P");
}

function layTinTucSheetChiTietMa(tenMa: string): void {
  const BASE_URL: string = "https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx";
  const mangDuLieuChinh: Array<[string, string, string, string]> = [];
  const QUERY_URL: string = `${BASE_URL}?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content: string = UrlFetchApp.fetch(QUERY_URL).getContentText();
  const $ = Cheerio.load(content);
  $("a").each(function (this: any) {
    const title: string = $(this).attr("title") ?? "";
    const link: string = `${BASE_URL}${$(this).attr("href") ?? ""}`;
    const date: string = $(this).siblings("span").text().substring(0, 10);
    mangDuLieuChinh.push([tenMa.toUpperCase(), title, link, date]);
  });
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "AH");
}

function layBaoCaoTaiChinh(): void {
  const mangDuLieuChinh: Array<[string, string, string, string]> = [];
  const tenMa: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "F1");
  const QUERY_URL: string = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
  const content: string = UrlFetchApp.fetch(QUERY_URL).getContentText();
  const $ = Cheerio.load(content);
  $("#divDocument>div>table>tbody>tr").each(function (this: any) {
    const title: string = $(this).children("td:nth-child(1)").text();
    const date: string = $(this).children("td:nth-child(2)").text();
    const link: string = $(this).children("td:nth-child(3)").children("a").attr("href") ?? "";
    mangDuLieuChinh.push([tenMa.toUpperCase(), title, date, link]);
  });
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh.slice(1, 11), SheetUtil.SHEET_DU_LIEU, 18, "AH");
}

interface ReportData {
  SourceName?: string;
  Title?: string;
  ReportTypeName?: string;
  LastUpdate?: string;
  Url?: string;
}

function layBaoCaoPhanTich(tenMa: string): void {
  const mangDuLieuChinh: [string, string, string, string, string][] = [];
  const url: string = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = SheetHttp.sendPostRequest(url); // Giả định về cấu trúc và kiểu dữ liệu của object

  const datas = Array.from(object.Data) as ReportData[];

  datas.forEach((element: { SourceName?: string, Title?: string, ReportTypeName?: string, LastUpdate?: string, Url?: string }) => {
    const SourceName: string = element.SourceName ?? "____";
    const Title: string = element.Title ?? "____";
    const ReportTypeName: string = element.ReportTypeName ?? "____";
    const LastUpdate: string = element.LastUpdate ?? "____";
    const Url: string = element.Url ?? "____";
    mangDuLieuChinh.push([SourceName, Title, ReportTypeName, LastUpdate, Url]);
  });

  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "AL");
}

function layThongTinCoDong(tenMa: string): void {
  const URL: string = `https://apipubaws.tcbs.com.vn/tcanalysis/v1/company/${tenMa}/large-share-holders`;

  const object = SheetHttp.sendRequest(URL); // Giả định về phương thức và kiểu trả về của sendRequest
  const mangDuLieuChinh: Array<[string, string, string]> = object.listShareHolder.map(
    ({ ticker, name, ownPercent }: { ticker: string; name: string; ownPercent: string }) => [ticker, name, ownPercent]
  );

  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "AD");
}


function layThongTongSoLuongCoPhieuDangNiemYet(tenMa: string): void {
  const market = 'HOSE';
  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/SecuritiesDetails?lookupRequest.market=${market}&lookupRequest.pageIndex=1&lookupRequest.pageSize=1000&lookupRequest.symbol=${tenMa}`;
  const token = SheetHttp.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: "get",
    headers: {
      "Authorization": token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    }
  }
  const object = SheetHttp.sendRequest(URL, OPTION); // Giả định về phương thức và kiểu trả về của sendRequest
  const mangDuLieuChinh: Array<[string]> = [];
  mangDuLieuChinh.push([object.data[0].RepeatedInfo[0].ListedShare]);
  SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_CHI_TIET_MA, 68, "G");
}

function batSukienSuaThongTinO(e: any) {
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (e.range.getA1Notation() === "F1" && sheet.getName() === SheetUtil.SHEET_CHI_TIET_MA) {
    layThongTinChiTietMa();
  }
}