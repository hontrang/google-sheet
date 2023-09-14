function layChiSoVnIndex() {
  const ngayHienTai = moment().format("YYYY-MM-DD");
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "A1");
  const thanhKhoanMoiNhat = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "D1");
  const url = "https://wgateway-iboard.ssi.com.vn/graphql";

  const query = "query indexQuery($indexIds: [String!]!) {indexRealtimeLatestByArray(indexIds: $indexIds) {indexID indexValue allQty allValue totalQtty totalValue advances declines nochanges ceiling floor change changePercent ratioChange __typename } }";
  const variables = `{"indexIds": ["VNINDEX" ]}`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const duLieuNhanVe = object.data.indexRealtimeLatestByArray[0];
  const thanhKhoan = duLieuNhanVe.totalValue * 1000000;
  console.log(duLieuNgayMoiNhat !== ngayHienTai);
  console.log(thanhKhoan !== thanhKhoanMoiNhat);
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, thanhKhoan]];
    SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_HOSE, 1, "A");
  } else {
    if (thanhKhoan !== thanhKhoanMoiNhat) {
      SheetUtility.chen1HangVaoDauSheet(SheetUtility.SHEET_HOSE);
      const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, thanhKhoan]];
      SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtility.SHEET_HOSE, 1, "A");
    }
  }

}

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  danhSachMa.shift();
  // const danhSachMa = ["LBM","LAF","L10","KSB","KPF"];
  const toDate = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_CAU_HINH, "B1");
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
  layThongTinPB(danhSachMa);
  layThongTinPE(danhSachMa);
  layThongTinRoomNuocNgoai(danhSachMa);
  layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa);
}

function layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay() {
  layGiaHangNgay();
  layKhoiLuongHangNgay();
  layKhoiNgoaiMuaHangNgay();
  layKhoiNgoaiBanHangNgay();
  SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G4");
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

  SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, "E", "A", tenMa);
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
  SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, "F", "A", tenMa);
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
  SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, "G", "A", tenMa);
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
  SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtility.SHEET_DU_LIEU, "I", "A", tenMa);
}

// function layGiaVaKhoiLuongTuanGanNhat() {
//   const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
//   danhSachMa.shift();
//   const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C4");
//   const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E4");
//   while (danhSachMa.length > 0) {
//     const MANG_PHU = danhSachMa.splice(0, 400);
//     const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=10000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${fromDate}~date:lte:${toDate}`
//     const object = SheetHttp.sendGetRequest(URL);

//     if (object?.data) {
//       MANG_PHU.map(tenMaMangPhu => {
//         const data = [];
//         let dataItems = [];
//         console.log(tenMaMangPhu);
//         object.data.map(item => {
//           if (item.code === tenMaMangPhu) {
//             dataItems.push(item);
//           }
//         });
//         const tenMa = tenMaMangPhu;
//         dataItems = dataItems.slice(0, 11).reverse();
//         const giaDongCua = dataItems.map(item => item.close * 1000);
//         const khoiLuong = dataItems.map(item => item.nmVolume);
//         data.push([...giaDongCua, ...khoiLuong]);
//         SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "AF", "A", tenMa);
//       });
//     } else {
//       data.push([Array(44).fill(0)]);
//       SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "AF", "A", tenMa);
//     }
//   }
//   SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G4");
// }

// function layKhoiNgoaiTuanGanNhat() {
//   const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
//   danhSachMa.shift();
//   const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C4");
//   const toDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "E4");
//   while (danhSachMa.length > 0) {
//     const MANG_PHU = danhSachMa.splice(0, 400);
//     const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${fromDate}~tradingDate:lte:${toDate}`;
//     const object = SheetHttp.sendGetRequest(URL);

//     if (object?.data) {
//       MANG_PHU.map(tenMaMangPhu => {
//         const data = [];
//         console.log(tenMaMangPhu);
//         let dataItems = [];
//         object.data.map(item => {
//           if (item.code === tenMaMangPhu) {
//             dataItems.push(item);
//           }
//         });
//         const tenMa = tenMaMangPhu;
//         dataItems = dataItems.slice(0, 11).reverse();
//         const foreignbuyvoltotal = dataItems.map(item => item.buyVol);
//         const foreignsellvoltotal = dataItems.map(item => item.sellVol);
//         data.push([...foreignbuyvoltotal, ...foreignsellvoltotal]);
//         SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "BB", "A", tenMa);
//       });
//     } else {
//       data.push([Array(22).fill(0)]);
//       SheetUtility.ghiDuLieuVaoDayTheoTenThamChieu(data, SheetUtility.SHEET_DU_LIEU, "BB", "A", tenMa);
//     }
//   }
//   SheetLog.logTime(SheetUtility.SHEET_CAU_HINH, "G4");
// }

// Hàm lấy giá tham chiếu từ một nguồn dữ liệu và ghi vào Google Sheets
function layGiaThamChieu() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  danhSachMa.shift();
  // const danhSachMa = ["LBM","LAF","L10","KSB","KPF"];
  const date = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_CAU_HINH, "B1");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 400);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=100&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`
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

function duLieuTam(date) {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  // const fromDate = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CAU_HINH, "C5");
  while (danhSachMa.length > 0) {
    const MANG_PHU = danhSachMa.splice(0, 400);
    const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`;
    const object = SheetHttp.sendGetRequest(URL);
    if (object?.data.length > 0) {
      const header = SheetUtility.layDuLieuTrongHang("Tam", 1);
      SheetUtility.chen1HangVaoSheet("Tam", 2);
      SheetUtility.ghiDuLieuVaoDay([[date]], "Tam", 2, 1);
      object.data.map(item => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.code) {
            SheetUtility.ghiDuLieuVaoDay([[item.close * 1000]], "Tam", 2, i + 1);
          }
        }
      });
    }
    console.log(date);
  }
}

function layKhoiNgoaiBanHangNgay() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const date = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "A1");
  const hangCuoi = SheetUtility.laySoHangTrongSheet(SheetUtility.SHEET_KHOI_NGOAI_BAN);
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_KHOI_NGOAI_BAN, hangCuoi, 1);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = SheetHttp.sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = SheetUtility.layDuLieuTrongHang(SheetUtility.SHEET_KHOI_NGOAI_BAN, 1);
        SheetUtility.ghiDuLieuVaoDay([["'" + date]], SheetUtility.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, 1);
        object.data.map(item => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetUtility.ghiDuLieuVaoDay([[item.sellVol]], SheetUtility.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log(date);
    }
  } else {
    console.log("done");
  }
}

function layKhoiNgoaiMuaHangNgay() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const date = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "A1");
  const hangCuoi = SheetUtility.laySoHangTrongSheet(SheetUtility.SHEET_KHOI_NGOAI_MUA);
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_KHOI_NGOAI_MUA, hangCuoi, 1);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object = SheetHttp.sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = SheetUtility.layDuLieuTrongHang(SheetUtility.SHEET_KHOI_NGOAI_MUA, 1);
        SheetUtility.ghiDuLieuVaoDay([["'" + date]], SheetUtility.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, 1);
        object.data.map(item => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetUtility.ghiDuLieuVaoDay([[item.buyVol]], SheetUtility.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log(date);
    }
  } else {
    console.log("done");
  }
}

function layKhoiLuongHangNgay() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const date = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "A1");
  const hangCuoi = SheetUtility.laySoHangTrongSheet(SheetUtility.SHEET_KHOI_LUONG);
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_KHOI_LUONG, hangCuoi, 1);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`
      const object = SheetHttp.sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = SheetUtility.layDuLieuTrongHang(SheetUtility.SHEET_KHOI_LUONG, 1);
        SheetUtility.ghiDuLieuVaoDay([["'" + date]], SheetUtility.SHEET_KHOI_LUONG, hangCuoi + 1, 1);
        object.data.map(item => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetUtility.ghiDuLieuVaoDay([[item.nmVolume]], SheetUtility.SHEET_KHOI_LUONG, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log(date);
    }
  } else {
    console.log("done");
  }
}

function layGiaHangNgay() {
  const danhSachMa = SheetUtility.layDuLieuTrongCot(SheetUtility.SHEET_DU_LIEU, "A");
  const date = SheetUtility.layDuLieuTrongOTheoTen(SheetUtility.SHEET_HOSE, "A1");
  const hangCuoi = SheetUtility.laySoHangTrongSheet(SheetUtility.SHEET_GIA);
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_GIA, hangCuoi, 1);
  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`;
      const object = SheetHttp.sendGetRequest(URL);
      if (object?.data.length > 0) {
        const header = SheetUtility.layDuLieuTrongHang(SheetUtility.SHEET_GIA, 1);
        SheetUtility.ghiDuLieuVaoDay([["'" + date]], SheetUtility.SHEET_GIA, hangCuoi + 1, 1);
        object.data.map(item => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetUtility.ghiDuLieuVaoDay([[item.close * 1000]], SheetUtility.SHEET_GIA, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log(date);
    }
  } else {
    console.log("done");
  }
}

// function doGet(e) {
//   if (e.parameter.chucnang === 'chiTietMa') {
//     layThongTinChiTietMa(e.parameter.ma);
//     return HtmlService.createHtmlOutput("Thành công");
//   } else {
//     return HtmlService.createHtmlOutput("Chức năng không đúng");
//   }
// }

function create2DArray(data) {
  const values = [];
  for (let i = 0; i < data.length; i++) {
    values.push([data[i]]);
  }
  return values;
}

function main() {
  const ngay = SheetUtility.layDuLieuTrongCot("TRUY VAN", "A");
  ngay.map(date => {
    duLieuTam(date);
  });
}