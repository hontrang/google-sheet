# ghi chú

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `1604570972` — đang hiện |
| **Vùng dùng** | `A1:J38`, 6 công thức |
| **Hằng số trong code** | không có |

## Chức năng

Sổ tay của người dùng — không tab nào và không code nào tham chiếu tới. Gồm bốn phần:

| Vùng | Nội dung |
| --- | --- |
| `A1:B…` | Link tham khảo (tên + URL), ví dụ fund flow sàn Singapore |
| `E1:J14` | **Các bước lựa chọn mã chứng khoán** — checklist: giá tăng dần 5 phiên; khối lượng giao dịch (tốt nhất > 1 triệu); thông tin cơ bản (tin xấu, huỷ niêm yết, "phốt", trái phiếu, cổ đông lớn gom, đơn vị kiểm toán, khối ngoại 2 tuần, có trong danh mục quỹ Dragon Capital không); cổ đông chính/ngoại có uy tín không; báo cáo phân tích gần nhất; cổ tức ≥ 5% |
| `E18:H24` | **Máy tính margin** nhỏ: vốn, tỷ lệ vay, số tiền vay, tỷ lệ ký quỹ tối thiểu, NAV, tỷ lệ ký quỹ hiện tại, tiền cần nộp thêm, tiền lãi vay theo lãi suất × số ngày |
| `A31:C33` | Ví dụ tính giá điều chỉnh sau chia cổ tức (FRT: 5% tiền mặt + 15% cổ phiếu) |
| `A37:C38` | Câu hỏi thường gặp (vì sao tổng KL trên bảng điện nhỏ hơn thực tế…) |

Checklist ở `E1:J14` chính là lý do tồn tại của nhiều khối trên [chi tiết mã](chi%20tiết%20mã.md) (cổ đông, kiểm toán, trái phiếu, báo cáo phân tích) và các tab [DC](DC.md), [Kiểm toán xấu](Kiểm%20toán%20xấu.md).
