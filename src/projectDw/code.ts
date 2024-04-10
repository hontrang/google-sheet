function layChiSoVnIndex(): void {
    const ngayHienTai: string = moment().format("YYYY-MM-DD");
    const duLieuNgayMoiNhat: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "A1");
    const thanhKhoanMoiNhat: number = parseFloat(SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "D1"));
    const url: string = "https://banggia.cafef.vn/stockhandler.ashx?index=true";

    const object = SheetHttp.sendPostRequest(url);
    const duLieuNhanVe = object[1];
    const thanhKhoan: number = parseFloat(duLieuNhanVe.value.replace(/,/g, '')) * 1000000000;
    console.log(duLieuNgayMoiNhat !== ngayHienTai);
    console.log(thanhKhoan !== thanhKhoanMoiNhat);
    if (duLieuNgayMoiNhat === ngayHienTai && thanhKhoan !== thanhKhoanMoiNhat) {
        const mangDuLieuChinh: [[string, number, number, number]] = [[ngayHienTai, duLieuNhanVe.index, duLieuNhanVe.percent / 100, thanhKhoan]];
        SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_HOSE, 1, "A");
    } else if (thanhKhoan !== thanhKhoanMoiNhat) {
        SheetUtil.chen1HangVaoDauSheet(SheetUtil.SHEET_HOSE);
        const mang_du_lieu_chinh: [[string, number, number, number]] = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.percent / 100, thanhKhoan]];
        SheetUtil.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, SheetUtil.SHEET_HOSE, 1, "A");
    } else {
        console.log("No action required");
    }
}

function layThongTinCoBan(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
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
    SheetLog.logTime(SheetUtil.SHEET_CAU_HINH, "G4");
}

function layGiaThamChieu(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
    const date: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_CAU_HINH, "B1");
    const data: [string, number][] = [];
    while (danhSachMa.length > 0) {
        const MANG_PHU: string[] = danhSachMa.splice(0, 400);
        const URL: string = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`;
        const object = SheetHttp.sendGetRequest(URL);

        if (object?.data.length > 0) {
            object.data.forEach((item: { code?: string, close?: number }) => {
                const code: string = item.code ?? "___";
                const close: number = item.close ?? 0;
                console.log(code);
                data.push([code, close * 1000]);
                SheetUtil.ghiDuLieuVaoDayTheoTen(data, SheetUtil.SHEET_DU_LIEU, 2, "K");
            });
        }

        // SheetLog.logTime(SheetUtil.SHEET_CAU_HINH, "D1");
    }
}


function layThongTinPB(danhSachMa: string[]): void {
    const QUERY_API: string = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
    const mangDuLieuChinh: number[][] = []; // Sử dụng kiểu mảng số để lưu giá trị

    for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
        const url: string = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
        const object: any = SheetHttp.sendGetRequest(url); // Sử dụng kiểu `any` cho đối tượng trả về; tốt hơn là định nghĩa kiểu cụ thể nếu có thể

        object.data.forEach((element: { value?: number }) => {
            const value: number = element.value ?? 0;
            mangDuLieuChinh.push([value]);
        });
    }

    SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "E");
}

function layThongTinPE(danhSachMa: string[]): void {
    const QUERY_API: string = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
    const mangDuLieuChinh: number[][] = []; // Kiểu mảng của mảng số để lưu trữ các giá trị

    for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
        const url: string = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
        const object: any = SheetHttp.sendGetRequest(url); // Sử dụng `any` cho object; tốt hơn nếu có interface cụ thể

        object.data.forEach((element: { value?: number }) => {
            const value: number = element.value ?? 0; // Sử dụng nullish coalescing operator (??) thay vì OR (||) để xử lý 0 một cách chính xác
            mangDuLieuChinh.push([value]);
        });
    }

    SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "F");
}

function layThongTinRoomNuocNgoai(danhSachMa: string[]): void {
    const QUERY_API: string = "https://finfo-api.vndirect.com.vn/v4";
    const mangDuLieuChinh: [number, number][] = []; // Kiểu mảng của cặp số

    for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
        const url: string = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
        const object: any = SheetHttp.sendGetRequest(url); // Khuyến khích định nghĩa kiểu cụ thể thay vì sử dụng `any`
        const jsonData = JSON.parse(object.data);
        jsonData.forEach((element: { totalRoom?: number, currentRoom?: number }) => {
            const totalRoom: number = element.totalRoom ?? 0;
            const currentRoom: number = element.totalRoom ?? 0;
            mangDuLieuChinh.push([totalRoom, currentRoom]);
        });
    }

    SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "G");
}

function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa: string[]): void {
    const QUERY_API: string = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
    const mangDuLieuChinh: number[][] = []; // Kiểu mảng của mảng số để lưu trữ các giá trị

    for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
        const url: string = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
        const object: any = SheetHttp.sendGetRequest(url);

        object.data.forEach((element: { value?: number }) => {
            const value: number = element.value ?? 0;
            mangDuLieuChinh.push([value]);
        });
    }

    SheetUtil.ghiDuLieuVaoDayTheoTen(mangDuLieuChinh, SheetUtil.SHEET_DU_LIEU, 2, "I");
}

function duLieuTam(): void {
    SheetUtil.layDuLieuTrongCot("TRUY VAN", "A").forEach((date: string) => {
        let danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");

        while (danhSachMa.length > 0) {
            const MANG_PHU: string[] = danhSachMa.splice(0, 400);
            const URL: string = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
            const object: any = SheetHttp.sendGetRequest(URL); // Khuyến khích định nghĩa kiểu cụ thể thay vì sử dụng `any`

            if (object?.data.length > 0) {
                const header: string[] = SheetUtil.layDuLieuTrongHang(SheetUtil.SHEET_KHOI_NGOAI_MUA, 1);
                const hangCuoi: number = SheetUtil.laySoHangTrongSheet(SheetUtil.SHEET_KHOI_NGOAI_MUA);

                SheetUtil.ghiDuLieuVaoDay([["'" + date]], SheetUtil.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, 1);

                object.data.forEach((item: any) => {
                    header.forEach((headerItem: string, i: number) => {
                        if (headerItem === item.code) {
                            SheetUtil.ghiDuLieuVaoDay([[item.buyVol]], SheetUtil.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, i + 1);
                        }
                    });
                });
            }
            console.log(date);
        }
    });
}

function layKhoiNgoaiBanHangNgay(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
    const date: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "A1");
    const hangCuoi: number = SheetUtil.laySoHangTrongSheet(SheetUtil.SHEET_KHOI_NGOAI_BAN);
    const duLieuNgayMoiNhat: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_KHOI_NGOAI_BAN, "A" + hangCuoi);

    if (duLieuNgayMoiNhat !== date) {
        while (danhSachMa.length > 0) {
            const MANG_PHU: string[] = danhSachMa.splice(0, 400);
            const URL: string = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
            const object: any = SheetHttp.sendGetRequest(URL); // Nên thay `any` bằng kiểu dữ liệu cụ thể nếu có thể.

            if (object?.data.length > 0) {
                const header: string[] = SheetUtil.layDuLieuTrongHang(SheetUtil.SHEET_KHOI_NGOAI_BAN, 1);
                SheetUtil.ghiDuLieuVaoDay([["'" + date]], SheetUtil.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, 1);
                object.data.map((item: any) => { // Nên thay `any` bằng kiểu dữ liệu cụ thể của `item`.
                    for (let i = 0; i < header.length; i++) {
                        if (header[i] === item.code) {
                            SheetUtil.ghiDuLieuVaoDay([[item.sellVol]], SheetUtil.SHEET_KHOI_NGOAI_BAN, hangCuoi + 1, i + 1);
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

function layKhoiNgoaiMuaHangNgay(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
    const date: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "A1");
    const hangCuoi: number = SheetUtil.laySoHangTrongSheet(SheetUtil.SHEET_KHOI_NGOAI_MUA);
    const duLieuNgayMoiNhat: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_KHOI_NGOAI_MUA, "A" + hangCuoi);

    if (duLieuNgayMoiNhat !== date) {
        while (danhSachMa.length > 0) {
            const MANG_PHU: string[] = danhSachMa.splice(0, 400);
            const URL: string = `https://finfo-api.vndirect.com.vn/v4/foreigns?size=10000&sort=tradingDate&q=code:${MANG_PHU.join(",")}~tradingDate:gte:${date}~tradingDate:lte:${date}`;
            const object: any = SheetHttp.sendGetRequest(URL); // Khuyến khích thay `any` bằng kiểu dữ liệu cụ thể nếu có thể.

            if (object?.data.length > 0) {
                const header: string[] = SheetUtil.layDuLieuTrongHang(SheetUtil.SHEET_KHOI_NGOAI_MUA, 1);
                SheetUtil.ghiDuLieuVaoDay([["'" + date]], SheetUtil.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, 1);
                object.data.map((item: any) => { // Khuyến khích thay `any` bằng kiểu dữ liệu cụ thể của `item`.
                    for (let i = 0; i < header.length; i++) {
                        if (header[i] === item.code) {
                            SheetUtil.ghiDuLieuVaoDay([[item.buyVol]], SheetUtil.SHEET_KHOI_NGOAI_MUA, hangCuoi + 1, i + 1);
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

function layKhoiLuongHangNgay(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
    const date: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "A1");
    const hangCuoi: number = SheetUtil.laySoHangTrongSheet(SheetUtil.SHEET_KHOI_LUONG);
    const duLieuNgayMoiNhat: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_KHOI_LUONG, "A" + hangCuoi);

    if (duLieuNgayMoiNhat !== date) {
        while (danhSachMa.length > 0) {
            const MANG_PHU: string[] = danhSachMa.splice(0, 400);
            const URL: string = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`;
            const object: any = SheetHttp.sendGetRequest(URL); // Nên thay thế `any` bằng kiểu dữ liệu cụ thể nếu có thể.

            if (object?.data.length > 0) {
                const header: string[] = SheetUtil.layDuLieuTrongHang(SheetUtil.SHEET_KHOI_LUONG, 1);
                SheetUtil.ghiDuLieuVaoDay([["'" + date]], SheetUtil.SHEET_KHOI_LUONG, hangCuoi + 1, 1);
                object.data.map((item: any) => {
                    for (let i = 0; i < header.length; i++) {
                        if (header[i] === item.code) {
                            SheetUtil.ghiDuLieuVaoDay([[item.nmVolume]], SheetUtil.SHEET_KHOI_LUONG, hangCuoi + 1, i + 1);
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

function layGiaHangNgay(): void {
    const danhSachMa: string[] = SheetUtil.layDuLieuTrongCot(SheetUtil.SHEET_DU_LIEU, "A");
    const date: string = SheetUtil.layDuLieuTrongOTheoTen(SheetUtil.SHEET_HOSE, "A1");
    const hangCuoi: number = SheetUtil.laySoHangTrongSheet(SheetUtil.SHEET_GIA);
    const duLieuNgayMoiNhat: string = SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_GIA, "A" + hangCuoi);

    if (duLieuNgayMoiNhat !== date) {
        while (danhSachMa.length > 0) {
            const MANG_PHU: string[] = danhSachMa.splice(0, 400);
            const URL: string = `https://finfo-api.vndirect.com.vn/v4/stock_prices?size=1000&sort=date&q=code:${MANG_PHU.join(",")}~date:gte:${date}~date:lte:${date}`;
            const object: any = SheetHttp.sendGetRequest(URL); // Nên thay `any` bằng kiểu dữ liệu cụ thể nếu có thể.

            if (object?.data.length > 0) {
                const header: string[] = SheetUtil.layDuLieuTrongHang(SheetUtil.SHEET_GIA, 1);
                SheetUtil.ghiDuLieuVaoDay([["'" + date]], SheetUtil.SHEET_GIA, hangCuoi + 1, 1);
                object.data.map((item: any) => { // Khuyến khích thay `any` bằng kiểu dữ liệu cụ thể của `item`.
                    for (let i = 0; i < header.length; i++) {
                        if (header[i] === item.code) {
                            SheetUtil.ghiDuLieuVaoDay([[item.close * 1000]], SheetUtil.SHEET_GIA, hangCuoi + 1, i + 1);
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

function taoMang2D<T>(data: T[]): T[][] {
    const values: T[][] = [];
    for (const element of data) {
        values.push([element]);
    }
    return values;
}


// function doGet(e) {
//   if (e.parameter.chucnang === 'chiTietMa') {
//     layThongTinChiTietMa(e.parameter.ma);
//     return HtmlService.createHtmlOutput("Thành công");
//   } else {
//     return HtmlService.createHtmlOutput("Chức năng không đúng");
//   }
// }


// function layThongTinPB(danhSachMa) {
//   const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
//   const mang_du_lieu_chinh = [];
//   for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
//     const url = `${QUERY_API}?order=reportDate&where=itemCode:51012&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
//     const object = SheetHttp.sendGetRequest(url);

//     object.data.forEach((element) => {
//       const value = element.value || 0;
//       mang_du_lieu_chinh.push([value]);
//     });
//   }

//   SheetUtil.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtil.SHEET_DU_LIEU, "E", "A", tenMa);
// }

// function layThongTinPE(danhSachMa) {
//   const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
//   const mang_du_lieu_chinh = [];
//   for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
//     const url = `${QUERY_API}?order=reportDate&where=itemCode:51006&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
//     const object = SheetHttp.sendGetRequest(url);

//     object.data.forEach((element) => {
//       const value = element.value || 0;
//       mang_du_lieu_chinh.push([value]);
//     });
//   }
//   SheetUtil.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtil.SHEET_DU_LIEU, "F", "A", tenMa);
// }

// function layThongTinRoomNuocNgoai(danhSachMa) {
//   const QUERY_API = "https://finfo-api.vndirect.com.vn/v4";
//   const mang_du_lieu_chinh = [];
//   for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
//     const url = `${QUERY_API}/ownership_foreigns/latest?order=reportedDate&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
//     const object = SheetHttp.sendGetRequest(url);

//     object.data.forEach((element) => {
//       mang_du_lieu_chinh.push([element.totalRoom, element.currentRoom]);
//     });
//   }
//   SheetUtil.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtil.SHEET_DU_LIEU, "G", "A", tenMa);
// }

// function layThongTinKhoiLuongTrungBinh10Ngay(danhSachMa) {
//   const QUERY_API = "https://api-finfo.vndirect.com.vn/v4/ratios/latest";
//   const mang_du_lieu_chinh = [];

//   for (let i = 0; i < danhSachMa.length; i += SheetUtil.KICH_THUOC_MANG_PHU) {
//     const url = `${QUERY_API}?order=reportDate&where=itemCode:51016&filter=code:${danhSachMa.slice(i, i + SheetUtil.KICH_THUOC_MANG_PHU).join(",")}`;
//     const object = SheetHttp.sendGetRequest(url);

//     object.data.forEach((element) => {
//       const value = element.value || 0;
//       mang_du_lieu_chinh.push([value]);
//     });
//   }
//   SheetUtil.ghiDuLieuVaoDayTheoTenThamChieu(mang_du_lieu_chinh, SheetUtil.SHEET_DU_LIEU, "I", "A", tenMa);
// }