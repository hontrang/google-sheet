// Khai bao bien toan cuc /////
let SHEET_THAM_CHIEU = "tham chiếu";
let SHEET_BANG_THONG_TIN = "bảng thông tin";
let SHEET_TIN_TUC = "tin tức";
let SHEET_DU_LIEU = "dữ liệu";
let SHEET_CHI_TIET_MA = "chi tiết mã";
let KICH_THUOC_MANG_PHU = 10;
let HEADER_MA = "Tên mã";
let HEADER_KHOI_LUONG = "Khối lượng"
let HEADER_CODE = "Tên sự kiện";
let HEADER_DATE = "Thời gian diễn ra sự kiện";
let HEADER_LINK = "Link sự kiện";
let mangPhu = new Array();
let mang_du_lieu_chinh = new Array();
let response;
let object;
let url;
let range;
let tenMa;


let OPTIONS = {
  method: 'GET',
  headers: {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  }
}

function hamThucThi() {
  getDataHose();
  laySuKienChungKhoan();
  layGiaVaKhoiLuongTheoMaChungKhoan();
  ganDuLieuVaoCot("A", "D")
}

function getDataHose() {
  url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const data = JSON.stringify({
    query: 'query stockRealtimesByGroup($group: String) {  stockRealtimesByGroup(group: $group) {    stockSymbol     matchedPrice }}',
    letiables: '{  "group": "HOSE"}'
  });
  let HEADER_CURRENT_PRICE = "giá hiện tại";
  let options = {
    "headers": {
      "Content-Type": "application/json; charset=utf-8"
    },
    "payload": data,
    "method": "POST"
  }
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  mang_du_lieu_chinh = object.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);

  // chèn header vào dữ liệu
  let headers = [HEADER_MA, HEADER_CURRENT_PRICE];
  mang_du_lieu_chinh.unshift(headers);

  // xoá trắng range
  sheet.getRange(1, 1, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  sheet.getRange(1, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length).setValues(mang_du_lieu_chinh);
}

function laySuKienChungKhoan() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CHI_TIET_MA);
  // lay dữ liệu ô F1
  let tenMa = sheet.getRange(1, 6).getValue();
  url = "https://finfo-api.vndirect.com.vn/v4/news?q=tagCodes:" + tenMa + "~newsGroup:disclosure,rights_disclosure~locale:VN&sort=newsDate:desc~newsTime:desc&size=10";
  response = UrlFetchApp.fetch(url, OPTIONS);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.map(({ newsTitle, newsUrl, newsDate }) => [newsTitle, newsUrl, newsDate]);
  let headers = [HEADER_CODE, HEADER_LINK, HEADER_DATE];

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  // xoá trắng range I3:K14
  sheet.getRange(3, 9, 11, 3).clearContent();
  range = sheet.getRange(3, 9, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);

  range.setValues(mang_du_lieu_chinh);
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  url = "https://msh-data.cafef.vn/graphql";
  let HEADER_TIME = "Thời điểm";
  let HEADER_CLOSED_PRICE = "Giá đóng cửa";
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CHI_TIET_MA);
  // lấy dữ liệu ô F1
  let tenMa = sheet.getRange(1, 6).getValue();
  // lấy dữ liệu ô F2
  let fromDate = sheet.getRange(2, 6).getValue();
  //lấy dữ liệu ô H2
  let toDate = sheet.getRange(2, 8).getValue();

  let query = 'query { tradingViewData(symbol: \"' + tenMa + '\", from: \"' + fromDate + '\",to: \"' + toDate + '\") { symbol close   volume    time   }  }  ';
  let options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
    },
    "payload": JSON.stringify({
      query
    })
  }
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.tradingViewData.map(({ time, open, close, volume }) => [new Date(time * 1000), close, volume]);
  let headers = [HEADER_TIME, HEADER_CLOSED_PRICE, HEADER_KHOI_LUONG];
  mang_du_lieu_chinh = mang_du_lieu_chinh.reverse();
  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);

  // xoá trắng range
  sheet.getRange(1, 1, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);

  // gán dữ liệu cho ô H1
  sheet.getRange(1, 8, 1, 1).setValue(object.data.tradingViewData[0].symbol);
  logTime(SHEET_CHI_TIET_MA, "J2");
}

function layThongTinPB() {
  let danhSachMa = layGiaTriTheoCot(SHEET_THAM_CHIEU, 6, 1);

  let HEADER2 = "P/B";
  let headers = [HEADER_MA, HEADER2];
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }

    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51012&filter=code:" + mangPhu.join(",");
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach(element => {
      try {
        mang_du_lieu_chinh.push(new Array(element.code, element.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(element.code, 0));
      }
    });
    mangPhu = [];
  }

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 10, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // xoá trắng range
  sheet.getRange(1, 10, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);
}

function layThongTinPE() {
  let danhSachMa = layGiaTriTheoCot(SHEET_THAM_CHIEU, 6, 1);
  let HEADER2 = "P/E";
  let headers = [HEADER_MA, HEADER2];
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);

  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51006&filter=code:" + mangPhu.join(',');
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach(data => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0));
      }
    });
    mangPhu = [];
  }

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 12, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // xoá trắng range
  sheet.getRange(1, 12, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);
}

function layThongTinRoomNuocNgoai() {
  let danhSachMa = layGiaTriTheoCot(SHEET_THAM_CHIEU, 6, 1);
  let HEADER2 = "Sở hữu tối đa room nước ngoài";
  let HEADER3 = "Room nước ngoài còn lại";
  let headers = [HEADER_MA, HEADER2, HEADER3];
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  let mang5Ma = new Array();
  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mang5Ma.push(danhSachMa.shift());
    }
    url = "https://finfo-api.vndirect.com.vn/v4/ownership_foreigns/latest?order=reportedDate&filter=code:" + mang5Ma.join(',');
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach(data => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.totalRoom, data.currentRoom));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0, 0));
      }
    });
    mang5Ma = [];
  }


  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 14, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // xoá trắng range
  sheet.getRange(1, 14, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  let danhSachMa = layGiaTriTheoCot(SHEET_THAM_CHIEU, 6, 1);
  let headers = [HEADER_MA, HEADER_KHOI_LUONG];
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);

  while (danhSachMa.length > 0) {
    for (let i = 0; i < KICH_THUOC_MANG_PHU; i++) {
      mangPhu.push(danhSachMa.shift());
    }
    url = "https://api-finfo.vndirect.com.vn/v4/ratios/latest?order=reportDate&where=itemCode:51016&filter=code:" + mangPhu.join(',');
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    object.data.forEach(data => {
      try {
        mang_du_lieu_chinh.push(new Array(data.code, data.value));
      } catch (e) {
        mang_du_lieu_chinh.push(new Array(data.code, 0));
      }
    });
    mangPhu = [];
  }

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 17, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // xoá trắng range
  sheet.getRange(1, 17, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);
}

// function layDanhSachMa(){
//   url = "https://msh-data.cafef.vn/graphql/";
//   let query = "{ HOSE { stocks(take: 3000) { items { symbol currentPrice } } } }";
//   let options = {"headers": {
//                             "Content-Type": "application/json"
//                           },
//                 "payload": JSON.stringify({query}),
//                 "method": "POST"
//               }
//   let response = UrlFetchApp.fetch(url, options);
//   let object = JSON.parse(response.getContentText());
//   let values = object.data.HOSE.stocks.items.map(({symbol}) => [symbol]);
//   return values;
// }

function layGiaTriTheoCot(activeSheet, rowIndex, columnIndex) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(activeSheet);
  range = sheet.getRange(rowIndex, columnIndex, sheet.getLastRow() - rowIndex + 1);
  // xoá phần tử rỗng trong mảng
  return range.getValues().filter(function (el) {
    return el != "";
  });
}

function layTinTuc() {
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_TIN_TUC);
  let sheetBangThongTin = SpreadsheetApp.getActive().getSheetByName(SHEET_BANG_THONG_TIN);
  // lay du lieu cot J sheet bảng thông tin
  let listTenMa = sheetBangThongTin.getRange("J:J").getValues();
  listTenMa = listTenMa.filter(String);
  
  listTenMa.reverse().pop();
  // lay dữ liệu ô F1
  listTenMa.forEach(tenMa => {
    url = "https://finfo-api.vndirect.com.vn/v4/news?q=tagCodes:" + tenMa + "~newsGroup:disclosure,rights_disclosure~locale:VN&sort=newsDate:desc~newsTime:desc&size=10";
    response = UrlFetchApp.fetch(url, OPTIONS);
    object = JSON.parse(response.getContentText());
    let values1 = object.data.map(({ newsTitle, newsUrl, newsDate }) => [tenMa, newsTitle, newsUrl, newsDate]);
    Array.prototype.push.apply(mang_du_lieu_chinh, values1);
  })

  let headers = [HEADER_MA, HEADER_CODE, HEADER_LINK, HEADER_DATE];

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  // xoá trắng range from A1
  sheet.clearContents();

  range = sheet.getRange(1, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);

  range.setValues(mang_du_lieu_chinh);
}

function layGiaTuanGanNhat() {
  // let danhSachMa = ["STB"];
  let danhSachMa = layGiaTriTheoCot(SHEET_THAM_CHIEU, 6, 1);
  let headers = [HEADER_MA, "ngày T-5", "ngày T-4", "ngày T-3", "ngày T-2", "ngày T-1", "Giá ngày T"];
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  let fromDate = sheet.getRange("AA11").getValue();
  let toDate = sheet.getRange("AB11").getValue();
  url = "https://msh-data.cafef.vn/graphql";
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    Logger.log(tenMa);
    let query = 'query { tradingViewData(symbol: \"' + tenMa + '\", from: \"' + fromDate + '\",to: \"' + toDate + '\") { symbol close } }';
    let options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      "payload": JSON.stringify({
        query
      })
    }
    response = UrlFetchApp.fetch(url, options);
    object = JSON.parse(response.getContentText());
    try {
    mangPhu = object.data.tradingViewData;
    let len = object.data.tradingViewData.length;
      mang_du_lieu_chinh.push(new Array(tenMa,
        mangPhu[len-6].close,
        mangPhu[len-5].close,
        mangPhu[len-4].close,
        mangPhu[len-3].close,
        mangPhu[len-2].close,
        mangPhu[len-1].close,
        ));
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA",0,0,0,0,0));
    }
  }

  // chèn header vào dữ liệu
  mang_du_lieu_chinh.unshift(headers);

  range = sheet.getRange(1, 29, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // xoá trắng range
  sheet.getRange(1, 29, sheet.getLastRow(), mang_du_lieu_chinh[0].length).clearContent();

  range.setValues(mang_du_lieu_chinh);

  // in thời điểm lấy dữ liệu hoàn tất
  logTime(SHEET_THAM_CHIEU, "P2");
}

function ganDuLieuVaoCot(cotThamChieu, cotDoiDuLieu) {
  let sheetDuLieu = SpreadsheetApp.getActive().getSheetByName(SHEET_DU_LIEU);
  sheetDuLieu.getRange(cotDoiDuLieu + ":" + cotDoiDuLieu).clearContent();
  let data = sheetDuLieu.getRange(cotThamChieu + "2").getValue();
  Logger.log(data);
}
