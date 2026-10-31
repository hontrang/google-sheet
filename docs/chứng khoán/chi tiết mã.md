# chi tiết mã

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `2036414777` — đang hiện, cố định dòng 1 |
| **Vùng dùng** | `A1:N84`, 71 công thức, 5 biểu đồ, 15 rule định dạng có điều kiện |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetChiTietMa` |

## Chức năng

Trang tra cứu sâu **một mã**: gõ mã vào `B1`, trigger `onEdit` tự chạy `layThongTinChiTietMa()` để cào giá 1 năm, tin tức, báo cáo tài chính, báo cáo phân tích, cổ đông lớn, cổ tức, beta/free float, đơn vị kiểm toán, trái phiếu… rồi vẽ lại biểu đồ giá so với VN-INDEX.

Code ghi dữ liệu thô vào [dữ liệu](dữ%20liệu.md) (cột `P:AQ`); tab này chủ yếu **hiển thị** lại bằng `OFFSET`/`XLOOKUP`.

## Bố cục

| Vùng | Nội dung | Nguồn |
| --- | --- | --- |
| `B1` | **Ô nhập mã** — sửa ô này kích hoạt toàn bộ luồng cập nhật | nhập tay |
| `D1`, `G1` | Tên công ty, ngành | `XLOOKUP` sang `dữ liệu!I` / `H` |
| `I1` | Trạng thái: `...` khi đang chạy, thời điểm khi xong | code |
| `A15:D53` | Lịch sử giá: ngày, giá, khối lượng (`A16 = OFFSET('dữ liệu'!P:R,1,0,38,3)`), `D` = % thay đổi | `dữ liệu!P:R` |
| `E15:I16` | Tổng room ngoại, room còn lại, KL trung bình, P/E, P/B | `XLOOKUP` sang [tham chiếu](tham%20chiếu.md) |
| `E17:I18` | Hệ số Beta (code ghi `E18`), vốn hoá, KL nước ngoài còn được mua, tổng số CP, khối ngoại ròng 4 tuần | code + công thức trên `dữ liệu!X18:Y37` |
| `E19:I20` | Free float (code ghi `E20`), tiền/nợ ngắn hạn, tiền/tài sản, tiền/trái phiếu, tiền/số CP | công thức trên khối BCTC ở `E66:E73` |
| `E21:H31` | 10 tin tức gần nhất | `OFFSET('dữ liệu'!AH1,0,0,11,4)` |
| `E33:H43` | 10 cổ đông lớn: tên, % sở hữu, thay đổi, quốc gia | `OFFSET('dữ liệu'!AD1,0,0,11,4)` |
| `I33:N43` | Trái phiếu liên quan: mã TP, ngày phát hành, đáo hạn, mệnh giá, khối lượng; `N` = mệnh giá × KL; `M33` = tổng | `OFFSET('dữ liệu'!S40,0,0,10,5)` |
| `E44:J54` | 10 báo cáo phân tích: nguồn, tiêu đề, giá mục tiêu, ngày, link, khuyến nghị | `OFFSET('dữ liệu'!AL1,0,0,11,6)` |
| `A55:D65` | 10 báo cáo tài chính (cafef) | `OFFSET('dữ liệu'!AH18,0,0,10,4)` |
| `E55:…` | Lịch sử cổ tức | `dữ liệu!AO17:AQ` |
| `A66:D75` | Đơn vị kiểm toán theo năm | `dữ liệu!AH28:AK` |
| `E66:E73` | Tiền & tương đương tiền, Tổng tài sản, Nợ ngắn hạn (3 quý gần nhất) | `dữ liệu!AL28:AN37` |
| `A84` | Ghi chú tự do | nhập tay |

Biểu đồ: *biểu đồ khối lượng gần đây*, *Biểu đồ giá & khối lượng* (nguồn `dữ liệu!P:R`), *Khối ngoại mua/bán* (`dữ liệu!W18:Y37`), *VN-INDEX (tỷ VNĐ)* (`dữ liệu!AD18:AE37`), và biểu đồ chính có tiêu đề `*<MÃ>* - <tên công ty>` do `ChartHelper.updateChart()` xoá-và-vẽ-lại mỗi lần đổi mã (nhận diện bằng tiêu đề bắt đầu bằng `*`).

## Code liên quan

| Hàm | Trigger | Ghi vào |
| --- | --- | --- |
| [`batSukienSuaThongTinO()`](../../src/projectMain/code.ts#L358) | onEdit | Gọi `layThongTinChiTietMa()` khi sửa đúng `chi tiết mã!B1` |
| [`layThongTinChiTietMa()`](../../src/projectMain/code.ts#L58) | — | Điều phối các hàm dưới, ghi `I1` |
| [`layGiaVaKhoiLuongTheoMaChungKhoan()`](../../src/projectMain/code.ts#L80) | — | `dữ liệu!P:R` (vndirect, khoảng `cấu hình!B17:B18`) |
| [`layBaoCaoPhanTich()`](../../src/projectMain/code.ts#L134) | — | `dữ liệu!AL2:AQ11` (simplize) |
| [`layTinTucSheetChiTietMa()`](../../src/projectMain/code.ts#L99) | — | `dữ liệu!AH2:AK` (cafef) |
| [`layBaoCaoTaiChinh()`](../../src/projectMain/code.ts#L117) | — | `dữ liệu!AH18:AK27` (cafef) |
| [`layThongTinCoDong()`](../../src/projectMain/code.ts#L158) | — | `dữ liệu!AD2:AG11` (simplize) |
| [`layThongTinCoTuc()`](../../src/projectMain/code.ts#L169) | — | `dữ liệu!AO18:AQ` (simplize) |
| [`layHeSoBetaVaFreeFloat()`](../../src/projectMain/code.ts#L195) | — | `chi tiết mã!E18`, `E20` (vndirect) |
| [`layDonViKiemToan()`](../../src/projectMain/code.ts#L214) | — | `dữ liệu!AH29:AK` (vndirect) |
| [`layChiTietBaoCaoTaiChinh()`](../../src/projectMain/code.ts#L226) | — | `dữ liệu!AL29:AN37` (vndirect) |
| [`ChartHelper.updateChart()`](../../src/utility/ChartUtil.ts) | — | Biểu đồ chính, trục Y theo `cấu hình!B3:B5`, `D4:D5` |
| [`layThongTinTraiPhieu()`](../../src/projectMain/code.ts#L308) | — | `dữ liệu!S40:W69` (vietstock) |
| [`vietstock()`](../../src/projectMain/code.ts#L263) / [`derivative()`](../../src/projectMain/code.ts#L274) | mỗi 5 phút | Đọc `B1`, ghi token/kết quả vào `cấu hình!B19:B20` |

## Lưu ý

Tab có nhiều lỗi định dạng và công thức đã ghi nhận — xem [sheet-audit-projectMain.md](../sheet-audit-projectMain.md): A1 (đã sửa), A2, A3 (`H20` `#DIV/0!`), B4–B8, B12–B14, C8–C10.
