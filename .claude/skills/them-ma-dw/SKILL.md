---
name: them-ma-dw
description: Thêm hoặc xoá mã chứng khoán do người dùng chỉ định ở 4 tab Giá, Khối Lượng, KN Mua, KN Bán trong Google Sheet "data warehouse" (projectDw) — thêm vào cột cuối dòng 1, hoặc xoá cả cột của mã. Dùng khi người dùng muốn thêm, chèn, bổ sung, hoặc xoá, gỡ, bỏ một hay vài mã cổ phiếu trong sheet DW / data warehouse.
---

# Thêm / xoá mã chứng khoán trong sheet data warehouse

Chỉ làm với **những mã người dùng đưa ra** — không tự đề xuất hay thêm/xoá mã khác (mỗi mã thêm vào làm sheet nặng hơn).

Phạm vi hiện tại: 4 tab ma trận **Giá, Khối Lượng, KN Mua, KN Bán**. Các bước khác (tab `dữ liệu`, bảng tên `J:K`, file chứng khoán) sẽ bổ sung sau — không làm trong skill này.

Vì sao mã phải có ở cả 4 tab: `layGiaHangNgay()` / `layKhoiLuongHangNgay()` / `layKhoiNgoaiMuaHangNgay()` / `layKhoiNgoaiBanHangNgay()` (15:50 hằng ngày) dò mã theo tên ở dòng 1 của từng tab; `layGiaHangNgay()` lấy danh sách mã từ **Khối Lượng** nhưng ghi vào **Giá**.

Hai script, cùng dùng Sheets API (`src/offline/sheetsApi.mjs`). Cả hai mặc định **chỉ đọc và in kế hoạch**; `--apply` mới ghi/xoá.

| Script | Làm gì |
| --- | --- |
| [them-ma.mjs](them-ma.mjs) | Ghi mã vào cột cuối dòng 1 (thêm cột nếu lưới đã đầy). Tab nào đã có mã thì bỏ qua. |
| [xoa-ma.mjs](xoa-ma.mjs) | Xoá **cả cột** của mã ở tab nào có mã đó. Lưu backup cột ra `backup-dw/` (gitignore) trước khi xoá. |

## Thêm mã

1. **Lấy danh sách mã từ người dùng.** Nếu người dùng chưa nêu mã cụ thể, hỏi — không đoán.
2. **Xem trước**: `node .claude/skills/them-ma-dw/them-ma.mjs <MA> [<MA> ...]`
   - Script tra từng mã trên vndirect rồi in kế hoạch: `✔` tab đã có, `+` sẽ ghi vào ô nào.
   - Dừng và báo người dùng nếu vndirect không biết mã, hoặc có cảnh báo `⚠️` không phải `STOCK` / `listed`.
   - Nếu có cảnh báo **sau cột OR**: file chứng khoán chỉ import 4 tab tới cột `OR`, nên mã đó sẽ không lên dashboard. Báo rõ cho người dùng trước khi ghi.
3. **Đưa kế hoạch cho người dùng và chờ xác nhận.**
4. **Ghi**: thêm `--apply`. Script đọc lại dòng 1 và in `✔`/`✘` — báo đúng kết quả đó.
5. **Nhắc**: cột mới chỉ có tiêu đề, **không có lịch sử**; dữ liệu bắt đầu có từ lần chạy 15:50 kế tiếp. Mã chưa có ở tab `dữ liệu` nên chưa hiện trên dashboard.

## Xoá mã

**Mục đích: gỡ mã vừa thêm để test luồng thêm** (cột chỉ có tiêu đề, chưa có lịch sử). Xoá cột là **mất toàn bộ lịch sử** của mã ở 4 tab, nên script có khoá an toàn:

- Cột **chưa có lịch sử** → xoá được.
- Cột **đã có dữ liệu** → script in `🔒 BỊ CHẶN` và không xoá gì (kể cả các mã khác trong cùng lệnh). Chỉ mở khoá bằng `--xoa-ca-lich-su=<MA>,<MA>` ghi đúng từng mã.
- **Không bao giờ tự thêm `--xoa-ca-lich-su`.** Chỉ thêm khi người dùng nói rõ trong chat rằng muốn xoá *cả lịch sử* của *đúng mã đó* — sau khi đã thấy số ô dữ liệu sẽ mất. Lời đồng ý chung chung ("ok", "xoá đi") cho một lệnh xoá không tính là đồng ý mất lịch sử.

Khôi phục được bằng file backup của script, hoặc bằng *File → Lịch sử phiên bản* của Google Sheets.

1. **Lấy danh sách mã từ người dùng** (không đoán).
2. **Xem trước**: `node .claude/skills/them-ma-dw/xoa-ma.mjs <MA> [<MA> ...]`
   Script in: cột sẽ xoá ở từng tab, **số ô dữ liệu sẽ mất** (hoặc "không có lịch sử"), mã nào bị chặn, và cảnh báo:
   - ô ở tab `TRUY VAN` trỏ thẳng vào cột đó → sẽ thành `#REF!`;
   - mã còn ở `dữ liệu!A` → dòng đó sẽ ra `#N/A` (skill không xoá dòng này);
   - file chứng khoán có thể còn tra mã này.
3. **Đưa toàn bộ kế hoạch + cảnh báo cho người dùng, chờ xác nhận rõ ràng** cho việc xoá.
4. **Xoá**: thêm `--apply`. Script lưu backup, đọc lại dòng 1 để chắc vị trí chưa đổi (đổi thì dừng, chưa xoá gì), xoá cột (phải sang trái trong mỗi tab), rồi kiểm tra lại. Báo người dùng đường dẫn file backup và kết quả `✔`/`✘`.

Sau khi xoá, các mã phía sau dịch trái một cột; không công thức nào lệch vì hàm GAS, tab `dữ liệu` và file chứng khoán đều tra theo tên mã, còn tham chiếu trực tiếp ở `TRUY VAN` được Google Sheets tự dịch (trừ ô trỏ đúng cột bị xoá).

## Lưu ý

- **Giới hạn cột OR**: file chứng khoán import `Giá!A1:OR1` (và tương tự cho 3 tab kia) — xem công thức ô `A1`, `A2` ở tab Giá / Khối Lượng / KN Mua / KN Bán của file chứng khoán. Hiện 4 tab dùng tới `OQ`, nên chỉ còn **một** chỗ (`OR`) cho mã mới lên được dashboard. Xoá bớt mã sẽ giải phóng chỗ.
- **Lỗi xác thực** (`invalid_grant`, thiếu scope `spreadsheets`): token ghi nằm ở `.sheets-token.json` (gitignore). Cấp lại bằng `node src/offline/sheetsApi.mjs login --rw` — mở trình duyệt, người dùng phải tự bấm Cho phép; chạy lệnh ở chế độ nền và đưa URL cho người dùng.
- Thứ tự cột giữa 4 tab không cần giống nhau (code dò theo tên); mã mới luôn được thêm vào **cuối** dòng 1.
