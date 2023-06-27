function getDataHose() {
  const url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const query = 'query stockRealtimesByGroup($group: String){stockRealtimesByGroup(group: $group){stockSymbol matchedPrice }}';
  const variables = '{"group":"HOSE"}';
  const response = SheetHttp.sendGraphQLRequest(url, query, variables);
  const stockData = response.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(stockData, SheetUtility.SHEET_DU_LIEU, 2, "A");
}

function layThongTinChiTietMa() {
  SheetLog.logStart(SheetUtility.SHEET_CHI_TIET_MA, "J2");
  layGiaVaKhoiLuongTheoMaChungKhoan();
  layBaoCaoPhanTich();
  layTinTucSheetChiTietMa();
  // layBaoCaoTaiChinh();
  layThongTinCoDong();
  SheetLog.logTime(SheetUtility.SHEET_CHI_TIET_MA, "J2");
}

function layTinTucSheetChiTietMa() {
  const BASE_URL = "https://s.cafef.vn";
  const NEWS_PATH = "/Ajax/Events_RelatedNews_New.aspx";
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
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

function layBaoCaoPhanTich() {
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = SheetHttp.sendPostRequest(url);
  const mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AL");
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F2");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "H2");

  const query = `query {tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") {symbol close volume  time }  }  `;
  const object = SheetHttp.sendGraphQLRequest(SheetHttp.URL_GRAPHQL_CAFEF, query, {});
  const mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close * 1000, volume]
  ).reverse();
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, 4);
}

function layThongTinCoDong() {
  const url = "https://finfo-iboard.ssi.com.vn/graphql";
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");

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

function layGiaVaKhoiLuongTuanGanNhat() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "C");
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_THAM_CHIEU, "T2");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_THAM_CHIEU, "U2");
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const query = "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) {stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }";

  const mang_du_lieu_chinh = danhSachMa.map((tenMa) => {
    console.log(tenMa);
    const variables = `{"symbol": "${tenMa}","offset": 1,"size": 30, "fromDate": "${fromDate}", "toDate": "${toDate}" }`;
    const object = SheetHttp.sendGraphQLRequest(url, query, variables);

    if (object?.data?.stockPrice?.dataList) {
      const dataItems = object.data.stockPrice.dataList.slice(0, 11).reverse();
      const closes = dataItems.map(item => item.closeprice);
      console.log(closes);
      const volumes = dataItems.map(item => item.totalmatchvol);
      const foreignbuyvoltotal = dataItems.map(item => item.foreignbuyvoltotal);
      const foreignsellvoltotal = dataItems.map(item => item.foreignsellvoltotal);

      return [tenMa, ...closes, ...volumes, ...foreignbuyvoltotal, ...foreignsellvoltotal];
    } else {
      return [tenMa, ...Array(44).fill(0)];
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AQ");
  SheetLog.logTime(SheetUtility.SHEET_THAM_CHIEU, "L2");
}