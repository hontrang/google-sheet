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
  url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const data = JSON.stringify({
    query:
      "query stockRealtimesByGroup($group: String) {  stockRealtimesByGroup(group: $group) {    stockSymbol     matchedPrice }}",
    variables: '{  "group": "HOSE"}',
  });
  let options = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  object = SheetHttp.sendRequest(url, options);
  mang_du_lieu_chinh = object.data.stockRealtimesByGroup.map(
    ({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]
  );
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 1);
}

function laySuKienChungKhoan() {
  // lay dữ liệu ô F1
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  mang_du_lieu_chinh = new Array();
  url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content = UrlFetchApp.fetch(url).getContentText();
  $ = Cheerio.load(content);
  $("a").each(function () {
    mang_du_lieu_chinh.push(new Array(tenMa.toUpperCase(), $(this).attr("title"), "https://s.cafef.vn" + $(this).attr("href"), $(this).siblings("span").text().substr(0, 10)));
  });
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 42);
}

function layBaoCaoPhanTich() {
  // lay dữ liệu ô F1
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  object = SheetHttp.sendRequest(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  });
  mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  SheetLog.logDebug(mang_du_lieu_chinh[0]);
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 46);
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
  // lấy dữ liệu ô F1
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  // lấy dữ liệu ô F2
  let fromDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 6);
  //lấy dữ liệu ô H2
  let toDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 8);

  let query = `query { tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") { symbol close   volume    time   }  }  `;
  let options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    payload: JSON.stringify({
      query,
    }),
  };
  object = SheetHttp.sendRequest(URL_GRAPHQL_CAFEF, options);
  mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close, volume]
  );
  mang_du_lieu_chinh = mang_du_lieu_chinh.reverse();
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 42);
  SheetUtility.ghiDuLieuVaoO(object.data.tradingViewData[0].symbol, SHEET_CHI_TIET_MA, 1, 8);
}

function layThongTinCoDong() {
  url = "https://finfo-iboard.ssi.com.vn/graphql";
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);

  const data = JSON.stringify({
    query:
      "query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) { shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }",
    variables: `{ "symbol": "${tenMa}", "size": 10, "offset": 1 }`,
  });
  let options = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  object = SheetHttp.sendRequest(url, options);
  mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 37);
}

function layThongTinPB() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51012&filter=code:" + mangPhu.join(",");
    object = SheetHttp.sendRequest(url, OPTIONS);
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push(new Array(element.code, element.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(element.code, 0));
      }
    });
    mangPhu = [];
  }
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 10);
}

function layThongTinPE() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51006&filter=code:" + mangPhu.join(",");
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
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 12);
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
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 14);
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
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 17);
}

function layTinTuc() {
  let sheetBangThongTin =
    SpreadsheetApp.getActive().getSheetByName(SHEET_BANG_THONG_TIN);
  // lay du lieu cot J sheet bảng thông tin
  let listTenMa = sheetBangThongTin.getRange("J:J").getValues();
  listTenMa = listTenMa.filter(String);

  listTenMa.reverse().pop();
  // lay dữ liệu ô F1
  listTenMa.forEach((tenMa) => {
    url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      mang_du_lieu_chinh.push(new Array(tenMa, $(this).attr("title"), "", "https://s.cafef.vn" + $(this).attr("href"), "", $(this).siblings("span").text().substr(0, 10)));
    });
  });
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_BANG_THONG_TIN, 40, 1);
}

function layGiaTuanGanNhat() {
  // let danhSachMa = ["STB"];
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  let fromDate = sheet.getRange("AA11").getValue();
  let toDate = sheet.getRange("AB11").getValue();
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    let query = `query { tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") { symbol close } }`;
    let options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      payload: JSON.stringify({
        query,
      }),
    };
    object = SheetHttp.sendRequest(URL_GRAPHQL_CAFEF, options);
    try {
      mangPhu = object.data.tradingViewData;
      let len = object.data.tradingViewData.length;
      mang_du_lieu_chinh.push(new Array(tenMa, mangPhu[len - 6].close, mangPhu[len - 5].close, mangPhu[len - 4].close, mangPhu[len - 3].close, mangPhu[len - 2].close, mangPhu[len - 1].close));
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0, 0, 0, 0, 0));
    }
  }
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 29);
  // in thời điểm lấy dữ liệu hoàn tất
  SheetLog.logTime(SHEET_THAM_CHIEU, "L2");
}

function layGiaThamChieu() {
  let danhSachMa = SheetUtility.layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_THAM_CHIEU);
  let fromDate = SheetDate.getDate(Date.parse(sheet.getRange("K1").getValue()));
  let toDate = SheetDate.getDate(Date.parse(fromDate) + 86400000);
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    let query = `query { tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") { symbol close } }`;
    let options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      payload: JSON.stringify({
        query,
      }),
    };
    object = SheetHttp.sendRequest(URL_GRAPHQL_CAFEF, options);
    try {
      mangPhu = object.data.tradingViewData;
      mang_du_lieu_chinh.push(
        new Array(mangPhu[0].symbol, mangPhu[0].close * 1000)
      );
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0));
    }
  }
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_THAM_CHIEU, 6, 1);

  // in thời điểm lấy dữ liệu hoàn tất
  SheetLog.logTime(SHEET_THAM_CHIEU, "K2");
}