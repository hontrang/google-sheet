//// Khai bao bien toan cuc /////
let SHEET_THAM_CHIEU = "tham chiếu";
let SHEET_BANG_THONG_TIN = "bảng thông tin";
let SHEET_DU_LIEU = "dữ liệu";
let SHEET_CHI_TIET_MA = "chi tiết mã";
let KICH_THUOC_MANG_PHU = 10;
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
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.stockRealtimesByGroup.map(
    ({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]
  );
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    1
  );
}

function laySuKienChungKhoan() {
  // lay dữ liệu ô F1
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  mang_du_lieu_chinh = new Array();
  url =
    "https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=" +
    tenMa +
    "&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2";
  const content = UrlFetchApp.fetch(url).getContentText();
  $ = Cheerio.load(content);
  $("a").each(function () {
    mang_du_lieu_chinh.push(new Array(tenMa.toUpperCase(), $(this).attr("title"), "https://s.cafef.vn" + $(this).attr("href"), $(this).siblings("span").text().substr(0, 10)));
  });
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    42
  );
}

function layBaoCaoPhanTich() {
  // lay dữ liệu ô F1
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  url = "https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:" + tenMa + "&pageIndex=1&pageSize=9"
  response = UrlFetchApp.fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    }
  });
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    46
  );
}

function layThongTinChiTietMa(){
  logStart(SHEET_CHI_TIET_MA, "J2");
  layGiaVaKhoiLuongTheoMaChungKhoan();
  layThongTinCoDong();
  layBaoCaoPhanTich();
  laySuKienChungKhoan();
  logTime(SHEET_CHI_TIET_MA, "J2");
  return "layThongTinChiTietMa completed";
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  url = "https://msh-data.cafef.vn/graphql";
  // lấy dữ liệu ô F1
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  // lấy dữ liệu ô F2
  let fromDate = layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 6);
  //lấy dữ liệu ô H2
  let toDate = layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 8);

  let query =
    'query { tradingViewData(symbol: "' +
    tenMa +
    '", from: "' +
    fromDate +
    '",to: "' +
    toDate +
    '") { symbol close   volume    time   }  }  ';
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
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close, volume]
  );
  mang_du_lieu_chinh = mang_du_lieu_chinh.reverse();

  ghiDuLieuVaoDayXoaCot(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    4
  );
  ghiDuLieuVaoO(
    object.data.tradingViewData[0].symbol,
    SHEET_CHI_TIET_MA,
    1,
    8
  );
}

function layThongTinCoDong() {
  url = "https://finfo-iboard.ssi.com.vn/graphql";
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);

  const data = JSON.stringify({
    query:
      "query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) { shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }",
    variables: '{ "symbol": \"' + tenMa + '\", "size": 10, "offset": 1 }',
  });
  let options = {
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
    payload: data,
    method: "POST",
  };
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );

  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    37
  );
}

function layThongTinPB() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }

    url =
      "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51012&filter=code:" +
      mangPhu.join(",");
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach((element) => {
      try {
        mang_du_lieu_chinh.push(new Array(element.code, element.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(element.code, 0));
      }
    });
    mangPhu = [];
  }
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    10
  );
}

function layThongTinPE() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url =
      "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51006&filter=code:" +
      mangPhu.join(",");
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach((data) => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0));
      }
    });
    mangPhu = [];
  }

  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    12
  );
}

function layThongTinRoomNuocNgoai() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let mang5Ma = new Array();
  mang_du_lieu_chinh = new Array();
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mang5Ma.push(danhSachMa.shift());
    }
    url =
      "https://finfo-api.vndirect.com.vn/v4/ownership_foreigns/latest?order=reportedDate&filter=code:" +
      mang5Ma.join(",");
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
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

  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    14
  );
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url =
      "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51016&filter=code:" +
      mangPhu.join(",");
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach((data) => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0));
      }
    });
    mangPhu = [];
  }

  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    17
  );
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
    url =
      "https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=" +
      tenMa +
      "&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2";
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      mang_du_lieu_chinh.push(new Array(tenMa, $(this).attr("title"), "", "https://s.cafef.vn" + $(this).attr("href"), "", $(this).siblings("span").text().substr(0, 10)));
    });
  });

  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_BANG_THONG_TIN,
    40,
    1
  );
}

function layGiaTuanGanNhat() {
  // let danhSachMa = ["STB"];
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  let fromDate = sheet.getRange("AA11").getValue();
  let toDate = sheet.getRange("AB11").getValue();
  url = "https://msh-data.cafef.vn/graphql";
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    Logger.log(tenMa);
    let query =
      'query { tradingViewData(symbol: "' +
      tenMa +
      '", from: "' +
      fromDate +
      '",to: "' +
      toDate +
      '") { symbol close } }';
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
    response = UrlFetchApp.fetch(url, options);
    object = JSON.parse(response.getContentText());
    try {
      mangPhu = object.data.tradingViewData;
      let len = object.data.tradingViewData.length;
      mang_du_lieu_chinh.push(
        new Array(
          tenMa,
          mangPhu[len - 6].close,
          mangPhu[len - 5].close,
          mangPhu[len - 4].close,
          mangPhu[len - 3].close,
          mangPhu[len - 2].close,
          mangPhu[len - 1].close
        )
      );
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0, 0, 0, 0, 0));
    }
  }
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_DU_LIEU,
    2,
    29
  );
  // in thời điểm lấy dữ liệu hoàn tất
  logTime(SHEET_THAM_CHIEU, "L2");
}

function layGiaThamChieu() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_THAM_CHIEU);
  let fromDate = getDate(Date.parse(sheet.getRange("K1").getValue()));
  let toDate = getDate(Date.parse(fromDate) + 86400000);
  url = "https://msh-data.cafef.vn/graphql";
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    Logger.log(tenMa + " " + danhSachMa.length);
    let query =
      'query { tradingViewData(symbol: "' +
      tenMa +
      '", from: "' +
      fromDate +
      '",to: "' +
      toDate +
      '") { symbol close } }';
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
    response = UrlFetchApp.fetch(url, options);
    object = JSON.parse(response.getContentText());
    try {
      mangPhu = object.data.tradingViewData;
      mang_du_lieu_chinh.push(
        new Array(mangPhu[0].symbol, mangPhu[0].close * 1000)
      );
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0));
    }
  }
  ghiDuLieuVaoDay(
    mang_du_lieu_chinh,
    SHEET_THAM_CHIEU,
    6,
    1
  );

  // in thời điểm lấy dữ liệu hoàn tất
  logTime(SHEET_THAM_CHIEU, "K2");
}