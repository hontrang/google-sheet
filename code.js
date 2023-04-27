function getDataHose() {
  const url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const query = 'query stockRealtimesByGroup($group: String){stockRealtimesByGroup(group: $group){stockSymbol matchedPrice }}';
  const variables = '{"group":"HOSE"}';
  const response = SheetHttp.sendGraphQLRequest(url, query, variables);
  const stockData = response.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(stockData, SheetUtility.SHEET_DU_LIEU, 2, "A");
  layChiSoVnIndex();
}

function layThongTinChiTietMa() {
  SheetLog.logStart(SheetUtility.SHEET_CHI_TIET_MA, "J2");
  layGiaVaKhoiLuongTheoMaChungKhoan();
  layThongTinCoDong();
  layBaoCaoPhanTich();
  layTinTucSheetChiTietMa();
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
  SheetUtility.ghiDuLieuVaoO(object.data.tradingViewData[0].symbol, SheetUtility.SHEET_CHI_TIET_MA, "H1");
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

function layThongTinPB() {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const danhSachMa = SheetUtility.layGiaTriTheoCot(SheetUtility.SHEET_DU_LIEU, 2, 3);
  const mang_du_lieu_chinh = [];
  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, SheetUtility.KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendGetRequest(url);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "J");
}

function layThongTinPE() {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const danhSachMa = SheetUtility.layGiaTriTheoCot(SheetUtility.SHEET_DU_LIEU, 2, 3);
  const mang_du_lieu_chinh = []
  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, SheetUtility.KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendGetRequest(url);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "L");
}

function layThongTinRoomNuocNgoai() {
  const danhSachMa = SheetUtility.layGiaTriTheoCot(SheetUtility.SHEET_DU_LIEU, 2, 3);
  const mang_du_lieu_chinh = [];
  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, SheetUtility.KICH_THUOC_MANG_PHU);
    const url = `https://finfo-api.vndirect.com.vn/v4/ownership_foreigns/latest?order=reportedDate&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendGetRequest(url);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.totalRoom, element.currentRoom]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "N");
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const danhSachMa = SheetUtility.layGiaTriTheoCot(SheetUtility.SHEET_DU_LIEU, 2, 3);
  const mang_du_lieu_chinh = [];
  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, SheetUtility.KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendGetRequest(url);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "Q");
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
  // const danhSachMa = ["HOT"];
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_THAM_CHIEU, "T2");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_THAM_CHIEU, "U2");
  const mang_du_lieu_chinh = [];
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  danhSachMa.forEach((tenMa) => {
    console.log(tenMa);
    const query = "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) {stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }";
    const variables = `{"symbol": "${tenMa}","offset": 1,"size": 30, "fromDate": "${fromDate}", "toDate": "${toDate}" }`;
    const object = SheetHttp.sendGraphQLRequest(url, query, variables);
    if (object?.data?.stockPrice?.dataList) {
      const closes = [tenMa];
      const volumes = [];
      const foreignbuyvoltotal = [];
      const foreignsellvoltotal = [];
      for (let i = 10; i >= 0; i--) {
        closes.push(object.data.stockPrice.dataList[i].closeprice);
        volumes.push(object.data.stockPrice.dataList[i].totalmatchvol);
        foreignbuyvoltotal.push(object.data.stockPrice.dataList[i].foreignbuyvoltotal);
        foreignsellvoltotal.push(object.data.stockPrice.dataList[i].foreignsellvoltotal);
      }
      closes.push(...volumes);
      closes.push(...foreignbuyvoltotal);
      closes.push(...foreignsellvoltotal);
      mang_du_lieu_chinh.push(closes);
    } else {
      mang_du_lieu_chinh.push([tenMa, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "AQ");
  SheetLog.logTime(SheetUtility.SHEET_THAM_CHIEU, "L2");
}

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "C");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_THAM_CHIEU, "I1");
  const fromDate = moment(toDate, "DD/MM/YYYY").subtract(1, "days").format("DD/MM/YYYY");
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const mang_du_lieu_chinh = danhSachMa.map(tenMa => {
    const query = "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) {stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }";
    const variables = `{"symbol": "${tenMa}", "offset": 1, "size": 1, "fromDate": "${fromDate}","toDate": "${toDate}"}`;
    const object = SheetHttp.sendGraphQLRequest(url, query, variables);

    try {
      let closeprice = object.data.stockPrice.dataList[0].closeprice;
      return [tenMa, closeprice];
    } catch (e) {
      return ["NA", 0];
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_THAM_CHIEU, 4, "A");
  SheetLog.logTime(SheetUtility.SHEET_THAM_CHIEU, "I2");
}

function layChiSoVnIndex() {
  const ngayHienTai = moment().format("YYYY-MM-DD");
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO("HOSE", "A1");
  const url = "https://wgateway-iboard.ssi.com.vn/graphql";

  const query = "query indexQuery($indexIds: [String!]!) {indexRealtimeLatestByArray(indexIds: $indexIds) {indexID indexValue allQty allValue totalQtty totalValue advances declines nochanges ceiling floor change changePercent ratioChange __typename } }";
  const variables = `{"indexIds": ["VNINDEX" ]}`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const duLieuNhanVe = object.data.indexRealtimeLatestByArray[0];

  if (duLieuNgayMoiNhat !== ngayHienTai) {
    SheetUtility.chen1HangVaoDauSheet("HOSE");
  }
  const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, duLieuNhanVe.totalValue * 1000000]];

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, "HOSE", 1, "A");
}