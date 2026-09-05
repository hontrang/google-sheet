# Audit công thức & hiển thị — Google Sheet "chứng khoán" (projectMain)

| | |
| --- | --- |
| **Sheet** | `chứng khoán` — `1GyxkNiyXantim6R6otooAGd6SLoR4J9Db8nCCx216og` |
| **GAS project** | `projectMain` (Dashboard), script ID `1LvkaTmx4VMfbQKfry_3KgU6aiQU1M7VaeMWlkp47rOyWbGel2D1GI30a` |
| **Ngày audit** | 2026-09-05 |
| **Phạm vi** | Toàn bộ 21 tab: công thức, number format, conditional formatting, data validation, 8 biểu đồ |
| **Cách thu thập** | Export live sheet sang `.xlsx` qua Google Drive rồi parse trực tiếp `xl/worksheets/*.xml`, `xl/charts/*.xml`, `xl/styles.xml` |

Ký hiệu mức độ:

| Mức | Nghĩa |
| --- | --- |
| **P0** | Đang cho kết quả sai / hiện lỗi trên màn hình ngay lúc này |
| **P1** | Nhãn, định dạng, tô màu lệch với dữ liệu — gây đọc nhầm |
| **P2** | Chạy đúng hôm nay nhưng mong manh, sẽ vỡ khi dữ liệu đổi |
| **P3** | Cấu hình / bảo mật / hiệu năng |

Cột **Xong**: sửa `` `[ ]` `` thành `` `[x]` `` khi đã xử lý xong mục đó.

---

## 1. Bảng tổng hợp

| ID | Mức | Vị trí | Lỗi | Cách sửa | Xong |
| --- | --- | --- | --- | --- | :-: |
| A1 | **P0** | [ChartUtil.ts:26-28](../src/utility/ChartUtil.ts#L26-L28) | Biểu đồ `chi tiết mã`: series "VNINDEX" vẽ nhầm **cột ngày** `'dữ liệu'!AD18:AD37` thay vì cột giá trị `AE`. Date serial ~46.220 nằm ngoài trục phụ 1600–2000 → đường VNINDEX **vô hình** | `addRange(sheet.getRange("'dữ liệu'!AE18:AE37"))` | `[x]` |
| A2 | **P0** | `chi tiết mã!D22`, `D23` | % thay đổi lệch 1 dòng: `D22=(B22-B24)/B24`, `D23=(B23-B25)/B25`, trong khi cả cột dùng `(Bn-Bn+1)/Bn+1`. `D23` sai cả dấu (−0,46% vs đúng +0,31%) | Copy `D24` đè lên `D22:D23` | `[ ]` |
| A3 | **P0** | `chi tiết mã!H20` | Hiện `#DIV/0!`. `=E67/'dữ liệu'!V39`, mà `V39=SUMPRODUCT(V40:V69,W40:W69)`=0 với mọi mã không có trái phiếu ([code.ts:349-352](../src/projectMain/code.ts#L349-L352) return sớm) | Bọc `IFERROR(...,"—")` | `[ ]` |
| A4 | **P0** | `bảng thông tin!G35:J68` | Khối "ngành / Tổng số mã / Số mã xanh / Tỷ trọng" **không có công thức nào** — toàn giá trị dán cứng, đang đóng băng `#DIV/0!` tại `I35/J35`, `I37/J37`, `I68/J68` và `#N/A` tại `I36/J36` | Dựng lại bằng `COUNTIFS` trên `tham chiếu` | `[ ]` |
| A5 | **P0** | `tham chiếu!O4:O400` | `=XLOOKUP(A:A, 'dữ liệu'!A:A, 'dữ liệu'!DF:DF)` — truyền cả cột làm search key, trong khi 14 cột lân cận đều dùng `A4`. Hiện chưa lộ vì cả 397 mã đều trả `KHÔNG` | Đổi thành `XLOOKUP($A4, …)` | `[ ]` |
| B1 | **P1** | `tham chiếu!K3` | Tiêu đề "giá ngày t" nhưng công thức là `(BC-BB)/BB` = **% thay đổi** (format `0.00%`, `O2:S2` đếm ngưỡng ±0,068) | Đổi nhãn thành "% thay đổi t/t-1" | `[ ]` |
| B2 | **P1** | `tham chiếu!L`, `M` | "trendline giá" / "trendline khối lượng" trỏ `'dữ liệu'!DC:DD` — chỉ có **6 / 5 ô** có dữ liệu trên ~880 dòng → **392/397 dòng trống** | Bổ sung nguồn `DC:DD` hoặc ẩn 2 cột | `[ ]` |
| B3 | **P1** | `tham chiếu!O3` | Nhãn hard-code "Giảm **9** phiên liên tiếp" trong khi ngưỡng thật lấy từ `cấu hình!B21` | Dùng `="Giảm "&'cấu hình'!B21&" phiên liên tiếp"` như `bảng thông tin!J1` | `[ ]` |
| B4 | **P1** | `chi tiết mã!D1` | Chứa tên công ty (text) nhưng number format `#,##0.00` | Đổi về `General` | `[ ]` |
| B5 | **P1** | `chi tiết mã!G22:G31` | Chứa URL tin tức nhưng format `M/d/yyyy` | Đổi về `General` | `[ ]` |
| B6 | **P1** | `chi tiết mã!H34:H40` | Chứa tên quốc gia (text) nhưng format `#,##0` | Đổi về `General` | `[ ]` |
| B7 | **P1** | `chi tiết mã!F34:F40` | Format `@` (plain text) → `% sở hữu` lưu thành **chuỗi** `'6.78'`, không sort/tính được. Gốc ở [code.ts:164](../src/projectMain/code.ts#L164) ép `${pctOfSharesOutHeld}` | Bỏ format `@`; ghi số thay vì template string | `[ ]` |
| B8 | **P1** | `chi tiết mã!E16` / `F16` | "Tổng room ngoại" `0%` vs "Room ngoại còn lại" `0.00%` — hai chỉ số cạnh nhau khác độ chính xác | Thống nhất `0.00%` | `[ ]` |
| B9 | **P1** | [code.ts:51](../src/projectMain/code.ts#L51) vs [code.ts:110](../src/projectMain/code.ts#L110) | Cùng nguồn tin cafef nhưng `bảng thông tin` ghi text `dd/MM/yyyy` (`31/08/2026`), `chi tiết mã` ghi date serial `yyyy-MM-dd` | Dùng `DateHelper.doiDinhDangNgay` ở cả hai chỗ | `[ ]` |
| B10 | **P1** | `bảng thông tin` CF | Rule `> 0` (xanh) phủ `C5:C7 … N21:N22`; rule `< 0` (đỏ) phủ **thêm** `C8, H8, A7:A8, F8, D8, G8, E8` → số âm tô đỏ nhưng số dương không tô xanh | Đồng bộ sqref của 3 rule | `[ ]` |
| B11 | **P1** | `bảng thông tin!G23:G31` CF | Rule `between 0.0001 and 0.065` **khai báo trùng 2 lần** | Xoá 1 rule | `[ ]` |
| B12 | **P1** | `chi tiết mã` CF | Dải chạy tới `D89:D1841`, `E290:E1841` trên sheet chỉ có nội dung tới ~dòng 76; còn áp nhầm lên `A77:A83` (cột A) | Thu về `D16:D53` | `[ ]` |
| B13 | **P1** | `chi tiết mã!G18` CF | colorScale mốc `0 / G18÷2 / G18` — max chính là ô đó nên **luôn ở màu cực đại**, gradient vô nghĩa | Lấy mốc từ ô cấu hình | `[ ]` |
| B14 | **P1** | `chi tiết mã!F18`, `H18` CF | colorScale chốt cứng `509725123094400` và `6396250200` — vốn hoá tăng là gradient bão hoà | Tham chiếu `MAX('tham chiếu'!…)` | `[ ]` |
| C1 | **P2** | `bảng thông tin!F23:F31` | Range trôi theo dòng: `F23`→`'lệnh'!B2:E597`, `F24`→`B2:E598`, `F25`→`B2:E599`, `F26`→`B2:E600`, `F27`→`B3:E601`… mỗi dòng lệch 1 (kéo công thức không khoá `$`). Thêm nữa `F23:F26` thiếu guard `if(A..="","",…)` mà `F28:F31` lại có | Khoá `$B$2:$E$605` cho cả cột | `[ ]` |
| C2 | **P2** | `bảng thông tin!A8` | `SUMIF($B$23:$B$33,"<>",$J$23:$J$31)` — criteria 11 dòng, sum range 9 dòng; Sheets tự nới thành `J23:J33` | Sửa thành `$J$23:$J$33` | `[ ]` |
| C3 | **P2** | `bảng thông tin!L23` | `COUNTA('dữ liệu'!T1:T16)` trong khi `A23`/`G23` dùng `T2:T16` → khối Lãi/lỗ tràn thừa 1 dòng | Đổi về `T2:T16` | `[ ]` |
| C4 | **P2** | `Giá!A2`, `Khối Lượng!A2`, `KN Mua!A2`, `KN Bán!A2` | Lỗi gõ `":OR2"&C1x` sinh range `Giá!A3308:OR23318` (dòng 23.318). Chỉ chạy đúng nhờ IMPORTRANGE tự cắt về cuối sheet nguồn | Sửa thành `":OR"&C1x` | `[ ]` |
| C5 | **P2** | `dữ liệu!C1`, `D1` vs `Giá!A1` | `dữ liệu` đọc `Giá!A1:OQ1`, `Giá` đọc `Giá!A1:OR1` — lệch 1 cột, một đường ống đang bỏ sót mã cuối | Thống nhất `OR` | `[ ]` |
| C6 | **P2** | `cấu hình!C11:C14` | C11/C12 = `COUNTA(...)`, C13/C14 = `COUNTA(...) + 1`; offset `-10` vs `-20`. Chạy đúng nhờ hằng số ngầm. C11=3318 vs C12=1982 — hai tab lẽ ra cùng số phiên mà chênh 1.336 dòng | Chuẩn hoá công thức + kiểm tra lệch dữ liệu DW | `[ ]` |
| C7 | **P2** | `cổ tức` — 26 data validation | Mỗi rule một start row khác nhau (`$A2:$A415`, `$A5:$A415`, `$A8:$A415`, `$A17:$A415`, `$A20:$A415`, `$A23:$A415`…) do kéo range tương đối → dropdown `A24` không chọn được mã ở `dữ liệu!A2:A22` | Khoá `$A$2:$A$415` cho tất cả | `[ ]` |
| C8 | **P2** | Bảng trái phiếu | Lệch số dòng 3 chiều: [code.ts:354](../src/projectMain/code.ts#L354) ghi tới 20 dòng (`pageSize: 20`), `chi tiết mã!I34=OFFSET('dữ liệu'!$S$40,0,0,**10**,5)` hiện 10 dòng, tổng `M33='dữ liệu'!V39` lại `SUMPRODUCT` trên **30** dòng (`V40:W69`) → tổng và bảng có thể không khớp | Thống nhất 20 dòng cả 3 nơi | `[ ]` |
| C9 | **P2** | `chi tiết mã!N34:N43` | `=sum(L34*M34)` — `SUM` bọc một tích đơn, thừa | `=L34*M34` | `[ ]` |
| C10 | **P2** | `chi tiết mã!A16` / `D` | `A16` spill 38 dòng (`A16:C53`) nhưng cột `D` chỉ có công thức tới `D52` → dòng cuối luôn trống | Kéo `D` tới `D53` | `[ ]` |
| C11 | **P2** | [code.ts:236-259](../src/projectMain/code.ts#L236-L259) | `layChiTietBaoCaoTaiChinh` ghi cứng dòng 29 / 32 / 35, mỗi block 3 dòng. Query `modelType:1,89,3,91` — nếu API trả >3 bản ghi cho một `itemCode`, block sau ghi đè header block kế tiếp | Chặn số dòng ghi mỗi block | `[ ]` |
| C12 | **P2** | [code.ts:220-224](../src/projectMain/code.ts#L220-L224) | `layDonViKiemToan` lặp không giới hạn (`size=100`) từ dòng 29, trong khi vùng hiển thị `A66=OFFSET(AH28,0,0,10,4)` chỉ 10 dòng | Cắt `slice(0, 10)` | `[ ]` |
| C13 | **P2** | [code.ts:347](../src/projectMain/code.ts#L347) | `${menhGia}`, `${khoiLuong}` ép số thành chuỗi trước khi ghi — cùng kiểu lỗi với B7, làm `SUMPRODUCT` ở `V39` trả 0 | Ghi số nguyên bản | `[ ]` |
| D1 | **P3** | [ChartUtil.ts:12-13](../src/utility/ChartUtil.ts#L12-L13) | `cấu hình!C4:C5` đã tính sẵn min/max VN-INDEX thật (1726,69 / 1853,08) nhưng code đọc `D4/D5` = 1600/2000 chốt cứng → 2 ô auto bị bỏ không | Đọc `C4`/`C5` | `[ ]` |
| D2 | **P3** | [ChartUtil.ts:25](../src/utility/ChartUtil.ts#L25) + `cấu hình!B3` | `setRange(LOW_MA − ABS_MA*2, HIGH_MA + ABS_MA*2)` với `B3=abs(B16-B17)*2 + B4/100` → biên độ trục Y phụ thuộc chênh lệch giá của **đúng 2 phiên gần nhất ×4**; gặp phiên trần/sàn là đường giá bị nén dẹt | Dùng `MIN*0.98 / MAX*1.02` | `[ ]` |
| D3 | **P3** | `dữ liệu!AS`, `BD`, `BO`, `CI` | ~397 dòng × 4 công thức `TRANSPOSE(OFFSET(INDIRECT(…ADDRESS(COUNTA(cả cột)…))))` ≈ **1.600 công thức volatile**, mỗi cái gọi `COUNTA` full-column 2 lần → tính lại toàn bộ mỗi khi workbook có bất kỳ thay đổi nào. Đây là nguyên nhân chính nếu sheet chậm | Thay `INDIRECT`/`OFFSET` bằng `INDEX`, cache `COUNTA` vào 1 ô | `[ ]` |
| D4 | **P3** | `cấu hình!B2`, `B7`, `B8`, `B9`, `B19` | Deployment ID, consumerID/consumerSecret SSI, số tài khoản, vietstock token lưu plain text trong sheet | Chuyển sang Script Properties | `[ ]` |
| D5 | **P3** | [code.ts:284](../src/projectMain/code.ts#L284), [289](../src/projectMain/code.ts#L289), [319](../src/projectMain/code.ts#L319), [324](../src/projectMain/code.ts#L324) | Cookie và `__RequestVerificationToken` của vietstock hard-code trong source, trong khi `vietstock()` đã có sẵn hàm lấy token động ghi vào `cấu hình!B19` | Đọc token từ `cấu hình!B19` | `[ ]` |

---

## 2. Thứ tự đề xuất xử lý

1. **A1** — sửa 1 dòng trong `ChartUtil.ts`, lấy lại đường VNINDEX trên biểu đồ chính.
2. **A2** — copy `D24` đè `D22:D23`.
3. **A3** — bọc `IFERROR` cho `H20`.
4. **A4** — dựng lại khối "ngành" ở `bảng thông tin` bằng công thức thay vì giá trị dán.
5. **A5** — sửa search key `tham chiếu!O`.
6. Nhóm **B** (nhãn/format) — rẻ, làm gộp một lượt.
7. **D3** — tối ưu hiệu năng, việc lớn nhất, nên làm riêng.

## 3. Ghi chú về phương pháp

- Google Sheets export sang `.xlsx` sẽ bọc các hàm riêng của GS (`IMPORTRANGE`, `QUERY`, `FILTER`, `ARRAYFORMULA`) trong `__xludf.DUMMYFUNCTION("<công thức gốc>")` kèm giá trị cache — công thức gốc vẫn đọc được nguyên văn.
- Ô **không có** thẻ `<f>` trong XML nghĩa là **không có công thức** trong sheet thật. Đây là căn cứ khẳng định A4.
- Biểu đồ do `ChartHelper.updateChart()` sinh ra là `xl/charts/chart8.xml` (title bắt đầu bằng `*`, khớp `ChartHelper.getChart`); tham chiếu series đọc trực tiếp từ `<c:ser>` nên A1 là kết luận chắc chắn, không phải suy đoán.
- Toàn bộ 21 tab đều đã được quét lỗi (`t="e"`): chỉ `bảng thông tin` có ô lỗi cứng; `chi tiết mã!H20` được Google export dưới dạng `t="str"` với giá trị `#DIV/0!`.
