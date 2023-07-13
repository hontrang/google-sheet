function getDataHose() {
  const url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const query = 'query stockRealtimesByGroup($group: String){stockRealtimesByGroup(group: $group){stockSymbol matchedPrice }}';
  const variables = '{"group":"HOSE"}';
  const response = SheetHttp.sendGraphQLRequest(url, query, variables);
  const stockData = response.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(stockData, SheetUtility.SHEET_DU_LIEU, 2, "A");
}

function layThongTinChiTietMa() {
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const KEY = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "B2");
  const QUERY_URL = `https://script.google.com/macros/s/${KEY}/exec?chucnang=chiTietMa&ma=${tenMa}`;
  UrlFetchApp.fetch(QUERY_URL);
  Logger.log("Hàm " + "layThongTinChiTietMa" + " chạy thành công");
}

function layBaoCaoTaiChinh() {
  const mang_du_lieu_chinh = [];
  const QUERY_API = "https://finfo-api.vndirect.com.vn/v4";
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const url = `${QUERY_API}/attachments?q=tagCodes:${tenMa}~type:FINANCIALSTATEMENT~locale:VN&sort=releasedDate:desc&size=10&page=1`;
  const object = SheetHttp.sendGetRequest(url);

  object.data.forEach((element) => {
    mang_du_lieu_chinh.push([element.title, "", element.fileLink, element.releasedDate]);
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 18, "AG");
}

function layTinTucSheetBangThongTin() {
  const mang_du_lieu_chinh = [];
  SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_BANG_THONG_TIN, "J").forEach((tenMa) => {
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
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_BANG_THONG_TIN, 40, "A");
}

function layThongTinChiTietMa() {
  SheetUtility.ghiDuLieuVaoO("...", SheetUtility.SHEET_CHI_TIET_MA, "J2");
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);
  layTinTucSheetChiTietMa(tenMa);
  // // layBaoCaoTaiChinh();
  layThongTinCoDong(tenMa);
  updateChart();
  SheetLog.logTime(SheetUtility.SHEET_CHI_TIET_MA, "J2");
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa) {
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F2");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "H2");

  const query = `query {tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") {symbol close volume  time }  }  `;
  const object = SheetHttp.sendGraphQLRequest(SheetHttp.URL_GRAPHQL_CAFEF, query, {});
  const mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close * 1000, volume]
  ).reverse();
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
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const query = "query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) {shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }";
  const variables = `{"symbol": "${tenMa}", "size": 10, "offset": 1 }`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AC");
}

function layBaoCaoTaiChinh() {
  const mang_du_lieu_chinh = [];
  const QUERY_API = "https://finfo-api.vndirect.com.vn/v4";
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const url = `${QUERY_API}/attachments?q=tagCodes:${tenMa}~type:FINANCIALSTATEMENT~locale:VN&sort=releasedDate:desc&size=10&page=1`;
  const object = SheetHttp.sendGetRequest(url);

  object.data.forEach((element) => {
    mang_du_lieu_chinh.push([element.title, "", element.fileLink, element.releasedDate]);
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 18, "AA");
}