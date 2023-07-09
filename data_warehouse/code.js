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

function layGiaVaKhoiLuongTuanGanNhat() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C4");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E4");
  const url = "https://finfo-iboard.ssi.com.vn/graphql";

  const query = "query stockPrice( $symbol: String! $size: Int $offset: Int $fromDate: String $toDate: String ) {stockPrice( symbol: $symbol size: $size offset: $offset fromDate: $fromDate toDate: $toDate ) }";
  let HANG_BAT_DAU = 2;
  danhSachMa.map((tenMa) => {
    const data = [];
    const variables = `{"symbol": "${tenMa}","offset": 1,"size": 30, "fromDate": "${fromDate}", "toDate": "${toDate}" }`;
    const object = SheetHttp.sendGraphQLRequest(url, query, variables);

    if (object?.data?.stockPrice?.dataList) {
      const dataItems = object.data.stockPrice.dataList.slice(0, 11).reverse();
      const closes = dataItems.map(item => item.closeprice);
      console.log(closes);
      const volumes = dataItems.map(item => item.totalmatchvol);
      const foreignbuyvoltotal = dataItems.map(item => item.foreignbuyvoltotal);
      const foreignsellvoltotal = dataItems.map(item => item.foreignsellvoltotal);

      data.push([tenMa, ...closes, ...volumes, ...foreignbuyvoltotal, ...foreignsellvoltotal]);
    } else {
      data.push( [tenMa, ...Array(44).fill(0)]);
    }
    SheetUtility.ghiDuLieuVaoDayTheoTen(data, SheetUtility.SHEET_DU_LIEU, HANG_BAT_DAU, "AE");
    HANG_BAT_DAU++;
  });

  SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G4");
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