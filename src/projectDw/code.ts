/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as Cheerio from 'cheerio';
interface RoomData {
  totalRoom?: number;
  currentRoom?: number;
}

interface VolumnData {
  value?: number;
}
function layChiSoVnIndex(): void {
  const ngayHienTai: string = DateHelper.layNgayHienTai('YYYY-MM-DD');
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const thanhKhoanMoiNhat: number = parseFloat(SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'D1'));
  const url = 'https://banggia.cafef.vn/stockhandler.ashx?index=true';

  const object = HttpHelper.sendPostRequest(url);
  const duLieuNhanVe = object[1];
  const thanhKhoan: number = parseFloat(duLieuNhanVe.value.replace(/,/g, '')) * 1000000000;
  if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
    const mangDuLieuChinh: [[string, number, number, number]] = [[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]];
    SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_HOSE, 1, 'A');
  } else if (thanhKhoan !== thanhKhoanMoiNhat) {
    SheetHelper.chen1HangVaoDauSheet(SheetHelper.SHEET_HOSE);
    const mangDuLieuChinh: [[string, number, number, number]] = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.percent / 100, thanhKhoan]];
    SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_HOSE, 1, 'A');
  } else {
    console.log('No action required');
  }
}

function layTyGiaUSDVND() {
  const tyGiaHomNay = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_DU_LIEU, 'K2');
  const ngayHomNay = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const duLieuNgayMoiNhat = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_TY_GIA, 'A1');
  if (duLieuNgayMoiNhat !== ngayHomNay) {
    SheetHelper.chen1HangVaoDauSheet(SheetHelper.SHEET_TY_GIA);
    SheetHelper.ghiDuLieuVaoDayTheoTen([[ngayHomNay, tyGiaHomNay]], SheetHelper.SHEET_TY_GIA, 1, 'A');
  } else {
    SheetHelper.ghiDuLieuVaoDayTheoTen([[ngayHomNay, tyGiaHomNay]], SheetHelper.SHEET_TY_GIA, 1, 'A');
  }
}

function layThongTinCoBan(): void {
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  layThongTinPB(danhSachMa);
  layThongTinPE(danhSachMa);
  layThongTinRoomNuocNgoai(danhSachMa);
  layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa);
}

// Hàm lấy giá, khối lượng và thông tin mua bán của khối ngoại hàng ngày
function layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay(): void {
  layGiaHangNgay();
  layKhoiLuongHangNgay();
  layKhoiNgoaiMuaHangNgay();
  layKhoiNgoaiBanHangNgay();
  LogHelper.logTime(SheetHelper.SHEET_CAU_HINH, 'G4');
}

function layGiaThamChieu(): void {
  const DEFAULT_FORMAT = 'YYYY-MM-DD';
  const DANH_SACH_MA: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  const date: string = DateHelper.changeFormatDate(SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_CAU_HINH, 'B1'), DEFAULT_FORMAT, 'DD/MM/YYYY');
  const market = 'HOSE';
  const mangDuLieuChinh: Array<[string]> = [];

  const URL = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${date}&lookupRequest.toDate=${date}&lookupRequest.market=${market}`;
  const token = HttpHelper.getToken();
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const object = HttpHelper.sendRequest(URL, OPTION);
  const datas = object.data;

  for (const element of DANH_SACH_MA) {
    const price = datas.filter((object: { Symbol: string }) => object.Symbol === element && object.Symbol.length == 3).map((object: { ClosePrice: number }) => object.ClosePrice);
    mangDuLieuChinh.push([price]);
  }
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'C');
}

function layThongTinPB(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const URL = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object: any = HttpHelper.sendGetRequest(URL);

    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, SheetHelper.SHEET_DU_LIEU, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SHEET_DU_LIEU, `E${vitri}:E${vitri}`);
    });
  }
}

function layThongTinPE(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object: any = HttpHelper.sendGetRequest(url);
    object.data.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, SheetHelper.SHEET_DU_LIEU, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SHEET_DU_LIEU, `F${vitri}:F${vitri}`);
    });
  }
}

function layThongTinRoomNuocNgoai(danhSachMa: string[]): void {
  const QUERY_API = 'https://finfo-api.vndirect.com.vn/v4';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object: any = HttpHelper.sendGetRequest(url);
    const datas = Array.from(object.data) as RoomData[];
    datas.forEach((element: { totalRoom?: number; currentRoom?: number; code?: string }) => {
      const totalRoom: number = element.totalRoom ?? 0;
      const currentRoom: number = element.currentRoom ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, SheetHelper.SHEET_DU_LIEU, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[totalRoom, currentRoom]], SheetHelper.SHEET_DU_LIEU, `G${vitri}:H${vitri}`);
    });
  }
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa: string[]): void {
  const QUERY_API = 'https://api-finfo.vndirect.com.vn/v4/ratios/latest';
  const duLieuCotThamChieu = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');

  for (let i = 0; i < danhSachMa.length; i += SheetHelper.KICH_THUOC_MANG_PHU) {
    const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetHelper.KICH_THUOC_MANG_PHU).join(',')}`;
    const object: any = HttpHelper.sendGetRequest(url);
    const datas = Array.from(object.data) as VolumnData[];
    datas.forEach((element: { code?: string; value?: number }) => {
      const value: number = element.value ?? 0;
      const tenMa: string = element.code ?? '_';
      const vitri = SheetHelper.layViTriCotThamChieu(tenMa, SheetHelper.SHEET_DU_LIEU, duLieuCotThamChieu, 2);
      SheetHelper.ghiDuLieuVaoDayTheoVung([[value]], SheetHelper.SHEET_DU_LIEU, `I${vitri}:I${vitri}`);
    });
  }
}

function duLieuTam(): void {
  SheetHelper.layDuLieuTrongCot('TRUY VAN', 'A').forEach((date: string) => {
    const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');

    while (danhSachMa.length > 0) {
      const MANG_PHU: string[] = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object: any = HttpHelper.sendGetRequest(URL); // Khuyến khích định nghĩa kiểu cụ thể thay vì sử dụng `any`

      if (object?.data.length > 0) {
        const header: string[] = SheetHelper.layDuLieuTrongHang(SheetHelper.SHEET_KHOI_NGOAI_MUA, 1);
        const hangCuoi: number = SheetHelper.laySoHangTrongSheet(SheetHelper.SHEET_KHOI_NGOAI_MUA);

        SheetHelper.ghiDuLieuVaoDay([["'" + date]], SheetHelper.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, 1);

        object.data.forEach((item: any) => {
          header.forEach((headerItem: string, i: number) => {
            if (headerItem === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.buyVol]], SheetHelper.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, i + 1);
            }
          });
        });
      }
      console.log(date);
    }
  });
}

function layKhoiNgoaiBanHangNgay(): void {
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  const date: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const hangCuoi: number = SheetHelper.laySoHangTrongSheet(SheetHelper.SHEET_KHOI_NGOAI_BAN);
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_KHOI_NGOAI_BAN, 'A' + hangCuoi);

  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU: string[] = danhSachMa.splice(0, 400);
      const URL = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object: any = HttpHelper.sendGetRequest(URL);

      if (object?.data.length > 0) {
        const header: string[] = SheetHelper.layDuLieuTrongHang(SheetHelper.SHEET_KHOI_NGOAI_BAN, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], SheetHelper.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          // Nên thay `any` bằng kiểu dữ liệu cụ thể của `item`.
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.sellVol]], SheetHelper.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối ngoại bán hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}

function layKhoiNgoaiMuaHangNgay(): void {
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  const date: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const hangCuoi: number = SheetHelper.laySoHangTrongSheet(SheetHelper.SHEET_KHOI_NGOAI_MUA);
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_KHOI_NGOAI_MUA, 'A' + hangCuoi);

  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU: string[] = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(',')}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
      const object: any = HttpHelper.sendGetRequest(url); // Khuyến khích thay `any` bằng kiểu dữ liệu cụ thể nếu có thể.

      if (object?.data.length > 0) {
        const header: string[] = SheetHelper.layDuLieuTrongHang(SheetHelper.SHEET_KHOI_NGOAI_MUA, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], SheetHelper.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          // Khuyến khích thay `any` bằng kiểu dữ liệu cụ thể của `item`.
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.buyVol]], SheetHelper.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối ngoại mua hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}

function layKhoiLuongHangNgay(): void {
  const danhSachMa: string[] = SheetHelper.layDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A');
  const date: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const hangCuoi: number = SheetHelper.laySoHangTrongSheet(SheetHelper.SHEET_KHOI_LUONG);
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_KHOI_LUONG, 'A' + hangCuoi);

  if (duLieuNgayMoiNhat !== date) {
    while (danhSachMa.length > 0) {
      const MANG_PHU: string[] = danhSachMa.splice(0, 400);
      const url = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(',')}~date:gte:${date}~date:lte:${date}`;
      const object: any = HttpHelper.sendGetRequest(url); // Nên thay thế `any` bằng kiểu dữ liệu cụ thể nếu có thể.

      if (object?.data.length > 0) {
        const header: string[] = SheetHelper.layDuLieuTrongHang(SheetHelper.SHEET_KHOI_LUONG, 1);
        SheetHelper.ghiDuLieuVaoDay([["'" + date]], SheetHelper.SHEET_KHOI_LUONG, hangCuoi + 1, 1);
        object.data.map((item: any) => {
          for (let i = 0; i < header.length; i++) {
            if (header[i] === item.code) {
              SheetHelper.ghiDuLieuVaoDay([[item.nmVolume]], SheetHelper.SHEET_KHOI_LUONG, hangCuoi + 1, i + 1);
            }
          }
        });
      }
      console.log('Lấy khối lượng hàng ngày thành công');
    }
  } else {
    console.log('done');
  }
}

function layGiaHangNgay(): void {
  const date: string = SheetHelper.layDuLieuTrongOTheoTen(SheetHelper.SHEET_HOSE, 'A1');
  const fromDate: string = DateHelper.changeFormatDate(date, 'YYYY-MM-DD', 'DD/MM/YYYY');
  const toDate: string = DateHelper.changeFormatDate(date, 'YYYY-MM-DD', 'DD/MM/YYYY');
  const hangCuoi: number = SheetHelper.laySoHangTrongSheet(SheetHelper.SHEET_GIA);
  const duLieuNgayMoiNhat: string = SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_GIA, 'A' + hangCuoi);
  const market = 'HOSE';

  if (duLieuNgayMoiNhat !== date) {
    const url = `https://fc-data.ssi.com.vn/api/v2/Market/DailyStockPrice?&lookupRequest.fromDate=${fromDate}&lookupRequest.toDate=${toDate}&lookupRequest.market=${market}`;
    const token = HttpHelper.getToken();
    const OPTION: URLFetchRequestOptions = {
      method: 'get',
      // eslint-disable-next-line @typescript-eslint/naming-convention
      headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
    };
    const object = HttpHelper.sendRequest(url, OPTION);

    if (object?.data.length > 0) {
      const header: string[] = SheetHelper.layDuLieuTrongHang(SheetHelper.SHEET_GIA, 1);
      SheetHelper.ghiDuLieuVaoDay([["'" + date]], SheetHelper.SHEET_GIA, hangCuoi + 1, 1);
      object.data.map((item: any) => {
        for (let i = 0; i < header.length; i++) {
          if (header[i] === item.Symbol) {
            SheetHelper.ghiDuLieuVaoDay([[item.ClosePrice]], SheetHelper.SHEET_GIA, hangCuoi + 1, i + 1);
          }
        }
      });
    }
    console.log('Lấy giá hàng ngày thành công');
  } else {
    console.log('done');
  }
}

function layDanhSachMa(): void {
  const market = 'HOSE';
  const pageIndex = 1;
  const pageSize = 1000;
  const token = HttpHelper.getToken();
  const url = `https://fc-data.ssi.com.vn/api/v2/Market/Securities?lookupRequest.market=${market}&lookupRequest.pageIndex=${pageIndex}&lookupRequest.pageSize=${pageSize}`;
  const OPTION: URLFetchRequestOptions = {
    method: 'get',
    // eslint-disable-next-line @typescript-eslint/naming-convention
    headers: { Authorization: token, 'Content-Type': 'application/json', Accept: 'application/json' }
  };
  const mangDuLieuChinh: Array<[string, string]> = [];
  const response = HttpHelper.sendRequest(url, OPTION);
  const datas = response.data;
  SheetHelper.xoaDuLieuTrongCot(SheetHelper.SHEET_DU_LIEU, 'A', 2, 2);
  for (const element of datas) {
    if (element.Symbol.length == 3) {
      mangDuLieuChinh.push([element.Symbol, element.StockName]);
    }
  }
  SheetHelper.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetHelper.SHEET_DU_LIEU, 2, 'A');
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_DANH_MUC_DC(url: string) {
  const result: any = [];
  const response = HttpHelper.sendRequest(url);
  const data = response.ffs_holding;
  console.log(response);
  data.forEach((element: any) => {
    const tenMa = element.stock ?? '_';
    const nhomNganh = element.sector_vi ?? '_';
    const tyLe = element.per_nav ?? '_';
    const sanGD = element.bourse_en ?? '_';
    const capNhatLuc = element.modified ?? '-';
    result.push([tenMa, nhomNganh, sanGD, tyLe, capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_THONG_TIN_TAI_SAN_DC(url: string) {
  const result: any = [];
  const response = HttpHelper.sendRequest(url);
  const data = response.ffs_asset;
  console.log(response);
  data.forEach((element: any) => {
    const tenTaiSan = element.name_vi ?? '_';
    const tyLe = element.weight ?? '_';
    const capNhatLuc = element.modified ?? '-';
    result.push([tenTaiSan, tyLe, capNhatLuc]);
  });
  return result;
}

/**
 * @customfunction
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
function LAY_SU_KIEN() {
  const result: any = [];
  const content = UrlFetchApp.fetch(`https://hontrang.github.io/tradingeconomics/`).getContentText();
  const $ = Cheerio.load(content);
  console.log($('table#calendar>thead.table-header,table#calendar>tbody').length);
  let date: string;
  $('table#calendar>thead.table-header,table#calendar>tbody').each(function (this: any) {
    if ($(this).attr('class') !== undefined) {
      date = $(this).find('tr>th:nth-child(1)').text().trim();
    } else {
      $(this)
        .children('tr')
        .each(function (this: any) {
          const timeValue = $(this).find('td:nth-child(1)').text().trim() || '00:00 AM';
          const thoigian = `${date} ${timeValue}`;
          const currency = $(this).find('td:nth-child(2) td.calendar-iso').text().trim();
          const name = $(this).find('td:nth-child(3)').text().trim();
          const actual = $(this).find('td:nth-child(4)').text().trim();
          const previous = $(this).find('td:nth-child(5)').text().trim();
          const forecast = $(this).find('td:nth-child(6)').text().trim();
          if (currency === 'US') {
            result.push([thoigian, currency, name, actual, forecast, previous]);
          }
        });
    }
  });
  return result;
}
