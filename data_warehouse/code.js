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

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "B1");
  const fromDate = moment(toDate, "DD/MM/YYYY").subtract(1, "days").format("DD/MM/YYYY");
  const url = "https://finfo-iboard.ssi.com.vn/graphql";
  const query = "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) {stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }";

  const mang_du_lieu_chinh = danhSachMa.map((tenMa) => {
    console.log(tenMa);
    const variables = `{"symbol": "${tenMa}","offset": 1,"size": 1, "fromDate": "${fromDate}", "toDate": "${toDate}" }`;
    const object = SheetHttp.sendGraphQLRequest(url, query, variables);

    if (object?.data?.stockPrice?.dataList) {
      try {
        const closeprice = object.data.stockPrice.dataList[0].closeprice;
        console.log(closeprice);
        return [tenMa, closeprice];
      } catch (e) {
        SheetLog.logDebug("unable to get data " + tenMa);
        return [tenMa, 0];
      }
    } else {
      return [tenMa, 0];
    }
  });

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "B");
  SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "D1");
}

function layThongTinCoBan() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  SheetUtility.ghiDuLieuVaoDayTheoTen(create2DArray(danhSachMa), SheetUtility.SHEET_DU_LIEU, 2, "D");
  layThongTinPB(danhSachMa);
  layThongTinPE(danhSachMa);
  layThongTinRoomNuocNgoai(danhSachMa);
  layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa);
}

function layThongTinPB(danhSachMa) {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const mang_du_lieu_chinh = [];
  for (let i = 0; i < danhSachMa.length; i += SheetUtility.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetUtility.KICH_THUOC_MANG_PHU).join(",")}`;
    const object = SheetHttp.sendGetRequest(url);

    object.data.forEach((element) => {
      const value = element.value || 0;
      mang_du_lieu_chinh.push([value]);
    });
  }

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "E");
}

function layThongTinPE(danhSachMa) {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const mang_du_lieu_chinh = [];
  for (let i = 0; i < danhSachMa.length; i += SheetUtility.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetUtility.KICH_THUOC_MANG_PHU).join(",")}`;
    const object = SheetHttp.sendGetRequest(url);

    object.data.forEach((element) => {
      const value = element.value || 0;
      mang_du_lieu_chinh.push([value]);
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "F");
}

function layThongTinRoomNuocNgoai(danhSachMa) {
  const QUERY_API = "https://finfo-api.vndirect.com.vn/v4";
  const mang_du_lieu_chinh = [];
  for (let i = 0; i < danhSachMa.length; i += SheetUtility.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetUtility.KICH_THUOC_MANG_PHU).join(",")}`;
    const object = SheetHttp.sendGetRequest(url);

    object.data.forEach((element) => {
      mang_du_lieu_chinh.push([element.totalRoom, element.currentRoom]);
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "G");
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa) {
  const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
  const mang_du_lieu_chinh = [];

  for (let i = 0; i < danhSachMa.length; i += SheetUtility.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetUtility.KICH_THUOC_MANG_PHU).join(",")}`;
    const object = SheetHttp.sendGetRequest(url);

    object.data.forEach((element) => {
      const value = element.value || 0;
      mang_du_lieu_chinh.push([value]);
    });
  }
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "I");
}

function layThongTinChiTietMa(tenMa) {
  SheetUtility.ghiDuLieuVaoO("...", SheetUtility.SHEET_CHI_TIET_MA, "G3");
  layGiaVaKhoiLuongTheoMaChungKhoan(tenMa);
  layBaoCaoPhanTich(tenMa);
  layTinTucSheetChiTietMa(tenMa);
  // layBaoCaoTaiChinh();
  layThongTinCoDong(tenMa);
  SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G3");
}

function layGiaVaKhoiLuongTheoMaChungKhoan(tenMa) {
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C3");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E3");

  const query = `query {tradingViewData(symbol: "${tenMa}", from: "${fromDate}",to: "${toDate}") {symbol close volume  time }  }  `;
  const object = SheetHttp.sendGraphQLRequest(SheetHttp.URL_GRAPHQL_CAFEF, query, {});
  const mang_du_lieu_chinh = object.data.tradingViewData.map(
    ({ time, close, volume }) => [new Date(time * 1000), close * 1000, volume]
  ).reverse();
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "J");
}

function layThongTinCoDong(tenMa) {
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const query = "query shareholders($symbol: String!, $size: Int, $offset: Int, $order: String, $orderBy: String, $type: String, $language: String) {shareholders( symbol: $symbol size: $size offset: $offset order: $order orderBy: $orderBy type: $type language: $language ) }";
  const variables = `{"symbol": "${tenMa}", "size": 10, "offset": 1 }`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const mang_du_lieu_chinh = object.data.shareholders.dataList.map(
    ({ ownershiptypecode, name, percentage, quantity, publicdate }) => [ownershiptypecode, name, percentage, quantity, publicdate.substr(0, 10)]
  );
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "M");
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

function layBaoCaoPhanTich(tenMa) {
  const url = `https://edocs.vietstock.vn/Home/Report_ReportAll_Paging?xml=Keyword:${tenMa}&pageIndex=1&pageSize=9`;
  const object = SheetHttp.sendPostRequest(url);
  const mang_du_lieu_chinh = object.Data.map(({ SourceName, Title, ReportTypeName, LastUpdate, Url }) => [SourceName, Title, ReportTypeName, LastUpdate, Url]);
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "V");
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
  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, 2, "R");
}

function doGet(e) {
  if (e.parameter.chucnang === 'chiTietMa') {
    layThongTinChiTietMa(e.parameter.ma);
    return HtmlService.createHtmlOutput("Thành công");
  } else {
    return HtmlService.createHtmlOutput("Chức năng không đúng");
  }
}

function create2DArray(data) {
  const values = [];
  for (let i = 0; i < data.length; i++) {
    values.push([data[i]]);
  }
  return values;
}