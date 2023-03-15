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
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  mang_du_lieu_chinh = new Array();
  url = `https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=${tenMa}&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2`;
  const content = UrlFetchApp.fetch(url).getContentText();
  $ = Cheerio.load(content);
  $("a").each(function () {
    mang_du_lieu_chinh.push(new Array(tenMa.toUpperCase(), $(this).attr("title"), "https://s.cafef.vn" + $(this).attr("href"), $(this).siblings("span").text().substr(0, 10)));
  });
  // ghi từ cột AP
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 42);
}

function layBaoCaoPhanTich() {
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  object = SheetHttp.sendRequest(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  });
  mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
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
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");
  let fromDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F2");
  let toDate = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "H2");

  url = "https://finfo-iboard.ssi.com.vn/graphql";

  const data = JSON.stringify({
    query:"query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) { stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }",
    variables: `{
      "symbol": "${tenMa}",
      "offset": 1,
      "size": 1000,
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
  mang_du_lieu_chinh = object.data.stockPrice.dataList.map(
    ({ tradingdate, closeprice, totalmatchvol }) => [tradingdate.slice(0, 10), closeprice, totalmatchvol]
  );
  // mang_du_lieu_chinh = mang_du_lieu_chinh.reverse();
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 4);
  SheetUtility.ghiDuLieuVaoO(object.data.stockPrice.dataList[0].symbol, SHEET_CHI_TIET_MA, "H1");
}

function layThongTinCoDong() {
  url = "https://finfo-iboard.ssi.com.vn/graphql";
  let tenMa = SheetUtility.layDuLieuTrongO(SHEET_CHI_TIET_MA, "F1");

  const data = JSON.stringify({
    query:"query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) { shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }",
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
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_BANG_THONG_TIN, 40, 1);
}

function layGiaTuanGanNhat() {
  // let danhSachMa = ["STB"];
  let danhSachMa = SheetUtility.layDuLieuTrongCot(SHEET_DU_LIEU, "C");
  let fromDate = SheetUtility.layDuLieuTrongO(SHEET_DU_LIEU, "AA11");
  let toDate = SheetUtility.layDuLieuTrongO(SHEET_DU_LIEU, "AB11");
  let mang_khoi_luong = new Array();
  // danhSachMa = danhSachMa.slice(1,10);
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift().toString();
    let query = `query { tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") { symbol close volume } }`;
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
      mang_khoi_luong.push(new Array(tenMa, mangPhu[len - 6].volume, mangPhu[len - 5].volume, mangPhu[len - 4].volume, mangPhu[len - 3].volume, mangPhu[len - 2].volume, mangPhu[len - 1].volume));
    } catch (e) {
      mang_du_lieu_chinh.push(new Array(tenMa, 0, 0, 0, 0, 0));
      mang_khoi_luong.push(new Array(tenMa, 0, 0, 0, 0, 0));
    }
  }
  SheetUtility.ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 29);
  SheetUtility.ghiDuLieuVaoDay(mang_khoi_luong, SHEET_DU_LIEU, 2, 51);
  // in thời điểm lấy dữ liệu hoàn tất
  SheetLog.logTime(SHEET_THAM_CHIEU, "L2");
}

function layGiaThamChieu() {
  let danhSachMa = SheetUtility.layDuLieuTrongCot(SHEET_DU_LIEU, "C");
  let fromDate = SheetDate.getDate(Date.parse(SheetUtility.layDuLieuTrongO(SHEET_THAM_CHIEU, "I1")));
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
  SheetLog.logTime(SHEET_THAM_CHIEU, "I2");
}