# cấu hình

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `393786342` — đang hiện |
| **Vùng dùng** | `A1:G21`, 4 công thức |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetCauHinh` |

## Chức năng

Tham số của data warehouse: ngày tham chiếu, khoảng ngày, định dạng ngày, URL API Dragon Capital và link ngược về file **chứng khoán**. Đồng thời là nơi code ghi dấu thời gian lần cập nhật cuối.

| Ô | Nhãn | Giá trị / công thức | Ai dùng |
| --- | --- | --- | --- |
| `B1` | Ngày tham chiếu | `QUERY(IMPORTRANGE(B13, "tham chiếu!I1:I1"))` — lấy từ [chứng khoán/tham chiếu](../chứng%20khoán/tham%20chiếu.md) | `layGiaThamChieu()` |
| `D1` | Ngày cập nhật | dấu thời gian cũ (2024), không code nào ghi nữa | `chứng khoán/tham chiếu!I2` import |
| `A3:G3` | Chi tiết mã | Từ ngày (cố định) / Đến ngày `TODAY()+1` / Cập nhật lúc | không còn dùng |
| `A4:G4` | Giá tuần gần nhất | Từ `TODAY()−20` (`C4`) / Đến `TODAY()+1` / **Cập nhật lúc** (`G4`) | `G4` ghi bởi `layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay()`; `C4`, `G4` được `chứng khoán/tham chiếu` import |
| `B6` | format date | `yyyy-MM-dd` | custom function DC, `layThongTinPhaiSinh()` |
| `B9` | Sự kiện | `*FOMC*` | không còn dùng (bảng thông tin dùng từ khoá ở file chứng khoán) |
| `B10` / `B11` | DCDS / DCDE | URL API `getLatestFundPortfolioData` cho từng quỹ | [DC](DC.md)`!A2`, `F2`, `I2`, `N2` |
| `B12` | Báo cáo DCDS | URL API báo cáo (code dùng URL viết cứng, không đọc ô này) | tham khảo |
| `B13` | docs chung khoan | URL file **chứng khoán** | `B1` |

## Lưu ý

`B1` của tab này lấy ngày tham chiếu từ file chứng khoán, trong khi file chứng khoán lại import `D1`/`C4`/`G4` từ đây — hai file `IMPORTRANGE` qua lại lẫn nhau. Cả hai chiều phải được cấp quyền truy cập ("Allow access") thì công thức mới tính được.
