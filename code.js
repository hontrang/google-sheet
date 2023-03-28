//// Khai bao hang so /////
const SHEET_THAM_CHIEU = "tham chiếu";
const SHEET_BANG_THONG_TIN = "bảng thông tin";
const SHEET_DU_LIEU = "dữ liệu";
const SHEET_CHI_TIET_MA = "chi tiết mã";
const SHEET_DEBUG = "debug";

const URL_GRAPHQL_CAFEF = "https://msh-data.cafef.vn/graphql";

const KICH_THUOC_MANG_PHU = 10;


//// Khai bao bien toan cuc /////
let mangPhu = new Array();
let mang_du_lieu_chinh = new Array();
let response;
let object;
let url;
let range;
let tenMa;

let OPTIONS = {
  method: "GET",
  headers: {
    "Content-Type": "application/json",
    Accept: "application/json",
  }
};

let OPTIONS_HTML = {
  method: "GET",
  headers: {
    "Content-Type": "text/html",
    Accept: "text/html",
  },
};

function getDataHose() {
  const API_URL = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const QUERY = `query stockRealtimesByGroup($group: String) {
    stockRealtimesByGroup(group: $group) {
      stockSymbol
      matchedPrice
    }
  }`;
  const variables = { group: "HOSE" };
  const response = SheetHttp.sendGraphQLRequest(API_URL, QUERY, variables);
  const stockData = response.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(stockData, SHEET_DU_LIEU, 2, "A");
}

function laySuKienChungKhoan() {
  const BASE_URL = "https://s.cafef.vn";
  const NEWS_PATH = "/Ajax/Events_RelatedNews_New.aspx";
  const tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  mang_du_lieu_chinh = [];
  const QUERY_URL = `${BASE_URL}${NEWS_PATH}?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content = UrlFetchApp.fetch(QUERY_URL).getContentText();
  $ = Cheerio.load(content);
  $("a").each(function () {
    const title = $(this).attr("title");
    const link = `${BASE_URL}${$(this).attr("href")}`;
    const date = $(this).siblings("span").text().substr(0, 10);
    mang_du_lieu_chinh.push([tenMa.toUpperCase(), title, link, date]);
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "AH");
}

function layBaoCaoPhanTich() {
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const OPTIONS = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  };
  object = SheetHttp.sendRequest(url, OPTIONS);
  mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "AL");
}

function layThongTinChiTietMa() {
  SheetLog.logStart(SHEET_CHI_TIET_MA, "J2");
  layGiaVaKhoiLuongTheoMaChungKhoan();
  layThongTinCoDong();
  layBaoCaoPhanTich();
  laySuKienChungKhoan();
  SheetLog.logTime(SHEET_CHI_TIET_MA, "J2");
  return "layThongTinChiTietMa completed";
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  let fromDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F2");
  let toDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "H2");

  url = "https://finfo-iboard.ssi.com.vn/graphql";

  const data = JSON.stringify({
    query: "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) { stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }",
    variables: `{
      "symbol": "${tenMa}",
      "offset": 1,
      "size": 1000,
      "fromDate": "${fromDate}",
      "toDate": "${toDate}"
    }`,
  });
  const OPTIONS = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  object = SheetHttp.sendRequest(url, OPTIONS);
  mang_du_lieu_chinh = object.data.stockPrice.dataList.map(
    ({ tradingdate, closeprice, totalmatchvol }) => [tradingdate.slice(0, 10), closeprice, totalmatchvol]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "D");
  SheetUtility.ghiDuLieuVaoO(object.data.stockPrice.dataList[0].symbol, SHEET_CHI_TIET_MA, "H1");
}

function layThongTinCoDong() {
  url = "https://finfo-iboard.ssi.com.vn/graphql";
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");

  const data = JSON.stringify({
    query: "query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) { shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }",
    variables: `{ "symbol": "${tenMa}", "size": 10, "offset": 1 }`,
  });
  const OPTIONS = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  object = SheetHttp.sendRequest(url, OPTIONS);
  mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "AC");
}

function layThongTinPB() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let mang_du_lieu_chinh = danhSachMa.map((ma) => {
    url = `https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51012&filter=code:${ma}`;
    let object = SheetHttp.sendRequest(url, OPTIONS);
    try {
      return [object.data[0].code, object.data[0].value];
    } catch (e) {
      return [ma, 0];
    }
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "J");
}

// code by chat-gpt
function layThongTinPE() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let mang_du_lieu_chinh = danhSachMa.map((ma) => {
    url = `https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51006&filter=code:${ma}`;
    let object = SheetHttp.sendRequest(url, OPTIONS);
    try {
      return [object.data[0].code, object.data[0].value];
    } catch (e) {
      return [ma, 0];
    }
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "L");
}

function layThongTinRoomNuocNgoai() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let mang5Ma = new Array();
  mang_du_lieu_chinh = new Array();
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mang5Ma.push(danhSachMa.shift());
    }
    url = "https://finfo-api.vndirect.com.vn/v4/ownership_foreigns/latest?order=reportedDate&filter=code:" + mang5Ma.join(",");
    object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((data) => {
      try {
        mang_du_lieu_chinh.push(
          new Array(data.code, data.totalRoom, data.currentRoom)
        );
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0, 0));
      }
    });
    mang5Ma = [];
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "N");
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51016&filter=code:" +
      mangPhu.join(",");
    object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((data) => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0));
      }
    });
    mangPhu = [];
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "Q");
}

function layTinTuc() {
  // lay du lieu cot J sheet bảng thông tin
  let listTenMa = SheetUtility.layDuLieuTrongCot(SHEET_BANG_THONG_TIN, "J");

  listTenMa.forEach((tenMa) => {
    url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      mang_du_lieu_chinh.push([tenMa, $(this).attr("title"), "", "https://s.cafef.vn" + $(this).attr("href"), "", $(this).siblings("span").text().substr(0, 10)]);
    });
  });
  // ghi dữ liệu từ ô A40
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_BANG_THONG_TIN, 40, "A");
}

function layGiaTuanGanNhat() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SHEET_DU_LIEU, "C");
  // const danhSachMa = ["YEG"];
  const fromDate = SheetUtility.layDuLieuTrongO(SHEET_THAM_CHIEU, "T2");
  const toDate = SheetUtility.layDuLieuTrongO(SHEET_THAM_CHIEU, "U2");
  const mang_du_lieu_chinh = [];
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  danhSachMa.forEach((tenMa) => {
    console.log(tenMa);
    const data = JSON.stringify({
      query: "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) { stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }",
      variables: `{
      "symbol": "${tenMa}",
      "offset": 1,
      "size": 30,
      "fromDate": "${fromDate}",
      "toDate": "${toDate}"
    }`,
    });
    let options = {
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
      payload: data,
      method: "POST",
    };
    object = SheetHttp.sendRequest(url, options);
    if (object?.data?.stockPrice?.dataList) {
      const closes = [tenMa];
      const volumes = [];
      const foreignbuyvoltotal = [];
      const foreignsellvoltotal = [];
      for (let i = 5; i >= 0; i--) {
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
      mang_du_lieu_chinh.push([tenMa, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "AQ");
  SheetLog.logTime(SHEET_THAM_CHIEU, "L2");
}

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SHEET_DU_LIEU, "C");
  const toDate = SheetUtility.layDuLieuTrongO(SHEET_THAM_CHIEU, "I1");
  const fromDate = moment(toDate, "DD/MM/YYYY").subtract(1, "days").format("DD/MM/YYYY");
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const mang_du_lieu_chinh = danhSachMa.map(tenMa => {
    const data = JSON.stringify({
      query: "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) { stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }",
      variables: `{
      "symbol": "${tenMa}",
      "offset": 1,
      "size": 1,
      "fromDate": "${fromDate}",
      "toDate": "${toDate}"
    }`,
    });
    let options = {
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
      payload: data,
      method: "POST",
    };
    object = SheetHttp.sendRequest(url, options);

    try {
      let closeprice = object.data.stockPrice.dataList[0].closeprice;
      return [tenMa, closeprice];
    } catch (e) {
      return ["NA", 0];
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_THAM_CHIEU, 4, "A");
  SheetLog.logTime(SHEET_THAM_CHIEU, "I2");
}
