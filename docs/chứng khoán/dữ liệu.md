# dữ liệu

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `1042544577` — đang hiện, cố định dòng 1 và cột `A` |
| **Vùng dùng** | `A1:DH398`, ~3.300 công thức, bộ lọc bật |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDuLieu` |

## Chức năng

Tab "kho trung gian" của Dashboard. Nó gom ba loại dữ liệu vào cùng một chỗ để các tab hiển thị chỉ cần `XLOOKUP`/`OFFSET`:

1. **Dữ liệu theo mã** (một mã mỗi dòng, `A:O` và `AS:DF`) — giá hiện tại, giá tham chiếu, thông tin cơ bản import từ data warehouse, chuỗi giá/khối lượng/khối ngoại 11–20 phiên gần nhất, sparkline, cờ giảm liên tiếp.
2. **Danh mục cá nhân** (`S:AC`) — tổng hợp từ [lệnh](lệnh.md).
3. **Vùng đệm cho [chi tiết mã](chi%20tiết%20mã.md)** (`P:R`, `W:AQ`) — code GAS ghi dữ liệu của mã đang xem vào đây.

Không nên sửa tay tab này, trừ cột `A`.

## Bố cục

### Dữ liệu theo mã

| Cột | Tiêu đề | Nguồn |
| --- | --- | --- |
| `A` | Danh sách mã | **nhập tay** — danh sách gốc mà cả `getDataHose()` lẫn [tham chiếu](tham%20chiếu.md) đều đọc |
| `B` | Giá hiện tại | **ghi bởi code** `getDataHose()` (vps.com.vn, ×1000) |
| `C:D` | Mã + giá ở ngày tham chiếu | `TRANSPOSE(QUERY(IMPORTRANGE(…"Giá!…")))` — lấy dòng ngày `cấu hình!B10` của [DW/Giá](../data%20warehouse/Giá.md) |
| `H` / `I` | Ngành / tên công ty | `QUERY(IMPORTRANGE(…"dữ liệu!A:I"))` từ [DW/dữ liệu](../data%20warehouse/dữ%20liệu.md), loại các mã `BBC|DAG|RDP|SVI` |
| `K:O` | P/B, P/E, room tối đa, room còn lại, KL TB 10 phiên | như trên |
| `AS:BC` | Giá 11 phiên gần nhất | `TRANSPOSE(OFFSET(INDIRECT(… 'Giá' …)))` trên tab [Giá](Giá.md) ẩn |
| `BD:BN` | Khối lượng 11 phiên | tương tự trên [Khối Lượng](Khối%20Lượng.md) |
| `BO:CH` | Khối ngoại mua 20 phiên | tương tự trên [KN Mua](KN%20Mua.md) |
| `CI:DB` | Khối ngoại bán 20 phiên | tương tự trên [KN Bán](KN%20Bán.md) |
| `DC` / `DD` | Trendline giá / khối lượng | `SPARKLINE` trên `AS:BC` / `BD:BN` |
| `DE` | GD khối ngoại ròng 20 phiên | `SUM(BO:CH) − SUM(CI:DB)` |
| `DF` | Giảm N phiên liên tiếp (`ĐẠT`/`KHÔNG`) | `SUMPRODUCT` so sánh từng cặp giá trong N cột trước `BC`, N = `cấu hình!B21` |

### Danh mục cá nhân

| Vùng | Nội dung |
| --- | --- |
| `S1:AC15` | stt, mã, số lượng, giá mua TB, giá hiện tại, tăng/giảm hôm qua, giá trị, tỷ trọng, lãi/lỗ, % lãi/lỗ, tổng vốn. Dòng 2 là `TIỀN MẶT`; `T3 = UNIQUE(FILTER('lệnh'!B:B, 'lệnh'!C:C="Mua"))` |
| `Y2` | Tiền mặt = vốn gốc + tiền bán − tiền mua + cổ tức − phí |
| `AC2` / `AC3` | Tổng giá trị / tổng số lượng CP |

### Vùng đệm cho chi tiết mã (ghi bởi code)

| Vùng | Nội dung | Hàm ghi |
| --- | --- | --- |
| `P1:R…` | ngày, giá, khối lượng 1 năm | [`layGiaVaKhoiLuongTheoMaChungKhoan()`](../../src/projectMain/code.ts#L80) |
| `S39:W69` | Trái phiếu (`V39` = tổng mệnh giá × KL) | [`layThongTinTraiPhieu()`](../../src/projectMain/code.ts#L308) |
| `W17:Y37` | Khối ngoại mua/bán 20 phiên của mã ở `chi tiết mã!B1` | công thức |
| `AD1:AG11` | Cổ đông lớn | [`layThongTinCoDong()`](../../src/projectMain/code.ts#L158) |
| `AD17:AG37` | VN-INDEX 20 phiên gần nhất | `QUERY(IMPORTRANGE(…"hose!A:D"), "limit 20")` |
| `AH1:AK11` | Tin tức | [`layTinTucSheetChiTietMa()`](../../src/projectMain/code.ts#L99) |
| `AH17:AK27` | Báo cáo tài chính | [`layBaoCaoTaiChinh()`](../../src/projectMain/code.ts#L117) |
| `AH28:AK…` | Đơn vị kiểm toán | [`layDonViKiemToan()`](../../src/projectMain/code.ts#L214) |
| `AL1:AQ11` | Báo cáo phân tích | [`layBaoCaoPhanTich()`](../../src/projectMain/code.ts#L134) |
| `AO17:AQ…` | Cổ tức | [`layThongTinCoTuc()`](../../src/projectMain/code.ts#L169) |
| `AL28:AN37` | Tiền, tổng tài sản, nợ ngắn hạn theo quý | [`layChiTietBaoCaoTaiChinh()`](../../src/projectMain/code.ts#L226) |

## Được dùng bởi

Hầu hết các tab của file: [tham chiếu](tham%20chiếu.md), [chi tiết mã](chi%20tiết%20mã.md), [bảng thông tin](bảng%20thông%20tin.md), [lệnh](lệnh.md) và các tab năm cũ, [cấu hình](cấu%20hình.md)`!C3:C5`.

## Lưu ý

- Cột `A` ở đây và cột `A` của DW `dữ liệu` phải đồng bộ thủ công. Khi bỏ mã huỷ niêm yết, phải sửa cả danh sách loại trừ `BBC|DAG|RDP|SVI` trong `H1`, `I1`, `K1`.
- ~1.600 công thức `INDIRECT`/`OFFSET` volatile ở `AS`, `BD`, `BO`, `CI` là nguyên nhân chính khiến sheet chậm — audit mục D3.
- Lệch cột `OQ`/`OR` giữa `C1:D1` và tab `Giá` — audit mục C5.
