function hamThucThi() {
  getDataHose();
}

function getDataHose() {
  url = "https://wgateway-iboard.ssi.com.vn/graphql/";
  const data = JSON.stringify({
    query: 'query stockRealtimesByGroup($group: String) {  stockRealtimesByGroup(group: $group) {    stockSymbol     matchedPrice }}',
    variables: '{  "group": "HOSE"}'
  });
  let options = {
    "headers": {
      "Content-Type": "application/json; charset=utf-8"
    },
    "payload": data,
    "method": "POST"
  }
  response = UrlFetchApp.fetch(url, options);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.stockRealtimesByGroup.map(({ stockSymbol, matchedPrice }) => [stockSymbol, matchedPrice]);
  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function laySuKienChungKhoan() {
  // lay dữ liệu ô F1
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  url = "https://finfo-api.vndirect.com.vn/v4/news?q=tagCodes:" + tenMa + "~newsGroup:disclosure,rights_disclosure~locale:VN&sort=newsDate:desc~newsTime:desc&size=10";
  response = UrlFetchApp.fetch(url, OPTIONS);
  object = JSON.parse(response.getContentText());
  mang_du_lieu_chinh = object.data.map(({ newsTitle, newsUrl, newsDate }) => [newsTitle, newsUrl, newsDate]);
  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_CHI_TIET_MA, 4, 9, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function layGiaVaKhoiLuongTheoMaChungKhoan() {
  url = "https://msh-data.cafef.vn/graphql";
  // lấy dữ liệu ô F1
  let tenMa = layDuLieuTrongO(SHEET_CHI_TIET_MA, 1, 6);
  // lấy dữ liệu ô F2
  let fromDate = layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 6);
  //lấy dữ liệu ô H2
  let toDate = layDuLieuTrongO(SHEET_CHI_TIET_MA, 2, 8);

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
  mang_du_lieu_chinh = object.data.tradingViewData.map(({ time, close, volume }) => [new Date(time * 1000), close, volume]);
  mang_du_lieu_chinh = mang_du_lieu_chinh.reverse();

  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_CHI_TIET_MA, 2, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  ghiDuLieuVaoO(object.data.tradingViewData[0].symbol, SHEET_CHI_TIET_MA, 1, 8, 1, 1);
  logTime(SHEET_CHI_TIET_MA, "J2");
}

function layThongTinPB() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
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
  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 10, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function layThongTinPE() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
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

  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 12, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function layThongTinRoomNuocNgoai() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
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

  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 14, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function layThongTinKhoiLuongTrungBinh10Ngay() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
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

  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 17, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
}

function layGiaTriTheoCot(activeSheet, rowIndex, columnIndex) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(activeSheet);
  range = sheet.getRange(rowIndex, columnIndex, sheet.getLastRow() - rowIndex + 1);
  // xoá phần tử rỗng trong mảng
  return range.getValues().filter(function (el) {
    return el != "";
  });
}

function layTinTuc() {
  let sheetBangThongTin = SpreadsheetApp.getActive().getSheetByName(SHEET_BANG_THONG_TIN);
  // lay du lieu cot J sheet bảng thông tin
  let listTenMa = sheetBangThongTin.getRange("J:J").getValues();
  listTenMa = listTenMa.filter(String);

  listTenMa.reverse().pop();
  // lay dữ liệu ô F1
  listTenMa.forEach(tenMa => {
    url = "https://s.cafef.vn/Ajax/Events_RelatedNews_New.aspx?symbol=" + tenMa + "&floorID=0&configID=0&PageIndex=1&PageSize=10&Type=2";
    const content = UrlFetchApp.fetch(url).getContentText();
    $ = Cheerio.load(content);
    $("a").each(function () {
      var title = $(this).attr("title");
      var url = "https://s.cafef.vn" + $(this).attr("href");
      var time = $(this).siblings("span").text();
      console.log(title, url, time);
      mang_du_lieu_chinh.push(new Array(tenMa, title, url, time));
    });
  });

  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_TIN_TUC, 2, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
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
        mangPhu[len - 6].close,
        mangPhu[len - 5].close,
        mangPhu[len - 4].close,
        mangPhu[len - 3].close,
        mangPhu[len - 2].close,
        mangPhu[len - 1].close,
      ));
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0, 0, 0, 0, 0));
    }
  }
  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_DU_LIEU, 2, 29, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);
  // in thời điểm lấy dữ liệu hoàn tất
  logTime(SHEET_THAM_CHIEU, "P2");
}

function layGiaThamChieu() {
  let danhSachMa = layGiaTriTheoCot(SHEET_DU_LIEU, 2, 3);
  let sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_THAM_CHIEU);
  let fromDate = getDate(Date.parse(sheet.getRange("K1").getValue()))
  let toDate = getDate(Date.parse(fromDate) + 86400000);
  url = "https://msh-data.cafef.vn/graphql";
  while (danhSachMa.length > 0) {
    tenMa = danhSachMa.shift();
    Logger.log(tenMa + " " + danhSachMa.length);
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
      mang_du_lieu_chinh.push(new Array(mangPhu[0].symbol,
        mangPhu[0].close * 1000));
    } catch (e) {
      mang_du_lieu_chinh.push(new Array("NA", 0));
    }
  }
  ghiDuLieuVaoDay(mang_du_lieu_chinh, SHEET_THAM_CHIEU, 6, 1, mang_du_lieu_chinh.length, mang_du_lieu_chinh[0].length);

  // in thời điểm lấy dữ liệu hoàn tất
  logTime(SHEET_THAM_CHIEU, "K2");
}