function getDataHose() {
  const DANH_SACH_MA = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const URL = `https://bgapidatafeed.vps.com.vn/getliststockdata/${DANH_SACH_MA.join(",")}`;
  const response = SheetHttp.sendGetRequest(URL);
  const stockData = response.map(({ sym, lastPrice }) => [sym, lastPrice * 1000]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(stockData, SheetUtility.SHEET_DU_LIEU, 2, "B");
}

function layTinTucSheetBangThongTin() {
  const mang_du_lieu_chinh = [];
  SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_CAU_HINH, "E").forEach((tenMa) => {
    const url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      const title = $(this).attr("title");
      const link = "https://s.cafef.vn" + $(this).attr("href");
      const date = $(this).siblings("span").text().substr(0, 10);
      mang_du_lieu_chinh.push([tenMa, title, "", link, "", date]);
    });
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_BANG_THONG_TIN, 35, "A");
}

function layThongTinChiTietMa() {
  SheetUtility.ghiDuLieuVaoO("...", SheetUtility.SHEET_CHI_TIET_MA, "J2");
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  // layBaoCaoPhanTich(tenMa);
  layTinTucSheetChiTietMa(tenMa);
  layBaoCaoTaiChinh();
  // layThongTinCoDong(tenMa);
  updateChart();
  SheetLog.logTime(SheetUtility.SHEET_CHI_TIET_MA, "J2");
  Logger.log("Hàm " + "layThongTinChiTietMa" + " chạy thành công");
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa) {
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F2");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "H2");

  const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?sort=date&q=code:${tenMa}~date:gte:${fromDate}~date:lte:${toDate}&size=1000`;
  const object = SheetHttp.sendGetRequest(URL);
  const mang_du_lieu_chinh = object.data.map(
    ({ date, close, nmVolume }) => [date, close * 1000, nmVolume]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "P");
}

function layBaoCaoPhanTich(tenMa) {
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = SheetHttp.sendPostRequest(url);
  const mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AL");
}

function layTinTucSheetChiTietMa(tenMa) {
  const BASE_URL = "https://s.cafef.vn";
  const NEWS_PATH = "/Ajax/Events_RelatedNews_New.aspx";
  const mang_du_lieu_chinh = [];
  const QUERY_URL = `${BASE_URL}${NEWS_PATH}?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content = UrlFetchApp.fetch(QUERY_URL).getContentText();
  $ = Cheerio.load(content);
  $("a").each(function () {
    const title = $(this).attr("title");
    const link = `${BASE_URL}${$(this).attr("href")}`;
    const date = $(this).siblings("span").text().substr(0, 10);
    mang_du_lieu_chinh.push([tenMa.toUpperCase(), title, link, date]);
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AH");
}

function layThongTinCoDong(tenMa) {
  const URL = `https://restv2.fireant.vn/symbols/${tenMa}/holders`;

  const object = SheetHttp.sendRequest(URL);
  const mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AC");
}

function layBaoCaoTaiChinh() {
  const mang_du_lieu_chinh = [];
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const QUERY_URL = `https://s.cafef.vn/Ajax/CongTy/BaoCaoTaiChinh.aspx?sym=${tenMa}`;
  const content = UrlFetchApp.fetch(QUERY_URL).getContentText();
  $ = Cheerio.load(content);
  $("#divDocument>div>table>tbody>tr").each(function () {
    const title = $(this).children("td:nth-child(1)").text();
    const date = $(this).children("td:nth-child(2)").text();
    const link = $(this).children("td:nth-child(3)").children("a").attr("href");
    mang_du_lieu_chinh.push([tenMa.toUpperCase(), title, date, link]);
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh.slice(1, 11), SheetUtility.SHEET_DU_LIEU, 18, "AH");
}

function batSukienSuaThongTinO(e) {
  if (e.range.getA1Notation() === "F1" && sheet.getName() === SheetUtility.SHEET_CHI_TIET_MA) {
    layThongTinChiTietMa();
    SheetUtility.ghiDuLieuVaoO("ok", SheetUtility.SHEET_CAU_HINH, "B6");
  } else {
    SheetUtility.ghiDuLieuVaoO("no ok", SheetUtility.SHEET_CAU_HINH, "B6");
  }
}