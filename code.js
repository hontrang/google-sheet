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

function layTinTucSheetChiTietMa() {
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
  layTinTucSheetChiTietMa();
  SheetLog.logTime(SHEET_CHI_TIET_MA, "J2");
  return "layThongTinChiTietMa completed";
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  const tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  const fromDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F2");
  const toDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "H2");

  const query = `query { tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") { symbol close   volume    time   }  }  `;

  const OPTIONS = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    payload: JSON.stringify({
      query,
    }),
  };
  object = SheetHttp.sendRequest(URL_GRAPHQL_CAFEF, OPTIONS);
  mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close * 1000, volume]
  ).reverse();
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 4);
  SheetUtility.ghiDuLieuVaoO(object.data.tradingViewData[0].symbol, SHEET_CHI_TIET_MA, "H1");
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
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);

  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "J");
}

function layThongTinPE() {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);

  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "L");
}

function layThongTinRoomNuocNgoai() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);

  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, KICH_THUOC_MANG_PHU);
    const url = `https://finfo-api.vndirect.com.vn/v4/ownership_foreigns/latest?order=reportedDate&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.totalRoom, element.currentRoom]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0, 0]);
      }
    });
  }

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "N");
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);

  while (danhSachMa.length > 0) {
    const mang_phu = danhSachMa.splice(0, KICH_THUOC_MANG_PHU);
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${mang_phu.join(",")}`;
    const object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push([element.code, element.value]);
      } catch (e) {
        mang_du_lieu_chinh.push([element.code, 0]);
      }
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, "Q");
}

function layTinTucSheetBangThongTin() {
  SheetUtility.layDuLieuTrongCot(SHEET_BANG_THONG_TIN, "J").forEach((tenMa) => {
    url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      const title = $(this).attr("title");
      const link = "https://s.cafef.vn" + $(this).attr("href");
      const date = $(this).siblings("span").text().substr(0, 10);
      mang_du_lieu_chinh.push([tenMa, title, "", link, "", date]);
    });
  });
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SHEET_BANG_THONG_TIN, 40, "A");
}

function layGiaVaKhoiLuongTuanGanNhat() {
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
    let OPTIONS = {
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
      payload: data,
      method: "POST",
    };
    object = SheetHttp.sendRequest(url, OPTIONS);
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
    let OPTIONS = {
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
      payload: data,
      method: "POST",
    };
    object = SheetHttp.sendRequest(url, OPTIONS);

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


function layChiSoVnIndex() {
  const ngayHienTai = moment().format("YYYY-MM-DD");
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO("HOSE", "A1");
  const url = "https://wgateway-iboard.ssi.com.vn/graphql";

  const data = JSON.stringify({
    query: "query indexQuery($indexIds: [String!]!) {   indexRealtimeLatestByArray(indexIds: $indexIds) {     indexID     indexValue     allQty     allValue     totalQtty     totalValue     advances     declines     nochanges     ceiling     floor     change     changePercent     ratioChange     __typename   } }",
    variables: `{
  "indexIds": [
    "VNINDEX"
  ]
}`,
  });
  let OPTIONS = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  object = SheetHttp.sendRequest(url, OPTIONS);
  const duLieuNhanVe = object.data.indexRealtimeLatestByArray[0];

  if (duLieuNgayMoiNhat !== ngayHienTai) {
    SheetUtility.chen1HangVaoDauSheet("HOSE");
  }
  const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, duLieuNhanVe.totalValue * 1000000]];

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, "HOSE", 1, "A");
}