function layChiSoVnIndex() {
  const ngayHienTai = moment().format("YYYY-MM-DD");
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO("HOSE", "A1");
  const thanhKhoanMoiNhat = SheetUtility.layDuLieuTrongO("HOSE", "D1");
  const url = "https://wgateway-iboard.ssi.com.vn/graphql";

  const query = "query indexQuery($indexIds: [String!]!) {indexRealtimeLatestByArray(indexIds: $indexIds) {indexID indexValue allQty allValue totalQtty totalValue advances declines nochanges ceiling floor change changePercent ratioChange __typename } }";
  const variables = `{"indexIds": ["VNINDEX" ]}`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const duLieuNhanVe = object.data.indexRealtimeLatestByArray[0];
  const thanhKhoan = duLieuNhanVe.totalValue * 1000000;
  console.log(duLieuNgayMoiNhat !== ngayHienTai);
  console.log(thanhKhoan !== thanhKhoanMoiNhat);
  if (duLieuNgayMoiNhat === ngayHienTai) {
    if (thanhKhoan !== thanhKhoanMoiNhat) {
      const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, thanhKhoan]];
      SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, "HOSE", 1, "A");
    }
  } else {
    if (thanhKhoan !== thanhKhoanMoiNhat) {
      SheetUtility.chen1HangVaoDauSheet("HOSE");
      const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, thanhKhoan]];
      SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, "HOSE", 1, "A");
    }
  }

}

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  // const danhSachMa = ["LBM","LAF","L10","KSB","KPF"];
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "B1");
  const fromDate = moment(toDate, "YYYY-MM-DD").subtract(1, "days").format("YYYY-MM-DD");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 5);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=100&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${fromDate}~date:lte:${toDate}`
    const object = SheetHttp.sendGetRequest(URL);

    if (object?.data) {
      MANG_PHU.map(tenMaMangPhu => {
        const data = [];
        console.log(tenMaMangPhu);
        let dataItems = [];
        object.data.map(item => {
          if (item.code === tenMaMangPhu) {
            dataItems.push(item);
          }
        });
        const tenMa = tenMaMangPhu;
        const giaDongCua = dataItems[0].close * 1000;
        // const ngay = dataItems.map(item => item.date);
        data.push([giaDongCua]);
        SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "B", "A", tenMa);
      });
    } else {
      data.push([Array(1).fill(0)]);
      SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "B", "A", tenMa);
    }
  }

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
  // const danhSachMa = ["YEG", "YBM", "VTO"];
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C4");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E4");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 5);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=200&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${fromDate}~date:lte:${toDate}`
    const object = SheetHttp.sendGetRequest(URL);

    if (object?.data) {
      MANG_PHU.map(tenMaMangPhu => {
        const data = [];
        console.log(tenMaMangPhu);
        let dataItems = [];
        object.data.map(item => {
          if (item.code === tenMaMangPhu) {
            dataItems.push(item);
          }
        });
        const tenMa = tenMaMangPhu;
        dataItems = dataItems.slice(0, 11).reverse();
        const giaDongCua = dataItems.map(item => item.close * 1000);
        const khoiLuong = dataItems.map(item => item.nmVolume);
        // const ngay = dataItems.map(item => item.date);
        data.push([...giaDongCua, ...khoiLuong]);
        SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "AF", "A", tenMa);
      });
    } else {
      data.push([Array(22).fill(0)]);
      SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "AF", "A", tenMa);
    }
  }
  SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G4");
}



// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  // const danhSachMa = ["LBM","LAF","L10","KSB","KPF"];
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "B1");
  const fromDate = moment(toDate, "YYYY-MM-DD").subtract(1, "days").format("YYYY-MM-DD");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 5);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=100&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${fromDate}~date:lte:${toDate}`
    const object = SheetHttp.sendGetRequest(URL);

    if (object?.data) {
      MANG_PHU.map(tenMaMangPhu => {
        const data = [];
        console.log(tenMaMangPhu);
        let dataItems = [];
        object.data.map(item => {
          if (item.code === tenMaMangPhu) {
            dataItems.push(item);
          }
        });
        const tenMa = tenMaMangPhu;
        const giaDongCua = dataItems[0].close * 1000;
        // const ngay = dataItems.map(item => item.date);
        data.push([giaDongCua]);
        SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "B", "A", tenMa);
      });
    } else {
      data.push([Array(1).fill(0)]);
      SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "B", "A", tenMa);
    }
  }

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

function layGiaHangNgay() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C5");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E5");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 400);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${fromDate}~date:lte:${toDate}`
    const object = SheetHttp.sendGetRequest(URL);

    if (object?.data) {
      const ngayHienTai = moment().format("YYYY-MM-DD");
      const header = SheetUtility.layDuLieuTrongHang("Giá", 1);
      SheetUtility.chen1HangVaoSheet("Giá", 2);
      SheetUtility.ghiDuLieuVaoDay([[ngayHienTai]], "Giá", 2, 1);
      object.data.map(item => {
        for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetUtility.ghiDuLieuVaoDay([[item.close * 1000]], "Giá", 2, i + 1);
            }
          }
      });
    }
  }
}

function layKhoiNgoaiTuanGanNhat() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  // const danhSachMa = ["YEG", "YBM", "VTO"];
  const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C4");
  const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E4");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 5);
    const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=200&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${fromDate}~tradingDate:lte:${toDate}`;
    const object = SheetHttp.sendGetRequest(URL);

    if (object?.data) {
      MANG_PHU.map(tenMaMangPhu => {
        const data = [];
        console.log(tenMaMangPhu);
        let dataItems = [];
        object.data.map(item => {
          if (item.code === tenMaMangPhu) {
            dataItems.push(item);
          }
        });
        const tenMa = tenMaMangPhu;
        dataItems = dataItems.slice(0, 11).reverse();
        const foreignbuyvoltotal = dataItems.map(item => item.buyVol);
        const foreignsellvoltotal = dataItems.map(item => item.sellVol);
        data.push([...foreignbuyvoltotal, ...foreignsellvoltotal]);
        SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "BB", "A", tenMa);
      });
    } else {
      data.push([Array(22).fill(0)]);
      SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "BB", "A", tenMa);
    }
  }
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