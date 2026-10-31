# Sự kiện

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `697191324` — **ẩn** |
| **Vùng dùng** | 1 công thức ở `A1`, spill ra 7 cột |
| **Hằng số trong code** | không có |

## Chức năng

Bản sao chỉ-đọc của lịch sự kiện kinh tế vĩ mô (chủ yếu Mỹ: CPI, FOMC, phát biểu của Fed, đấu giá trái phiếu…), kéo nguyên từ tab [Sự kiện của data warehouse](../data%20warehouse/Sự%20kiện.md):

```
A1 = QUERY(IMPORTRANGE('cấu hình'!B1, "Sự kiện!A:G"), "select *")
```

Cột: ngày (giờ Việt Nam, dạng `Thứ Ba 2026/10/13 07:15 CH`), quốc gia, tên sự kiện, kết quả, trước đó, nhất trí, dự đoán.

## Được dùng bởi

[bảng thông tin](bảng%20thông%20tin.md)`!F8:G8` — `XLOOKUP` wildcard tìm sự kiện đầu tiên khớp từ khoá ở `cấu hình!B15` / `B16` (mặc định `*FOMC*`, `*Fed Interest Rate*`) và ghép "tên - ngày - DĐ: dự đoán - KQ: kết quả".

## Lưu ý

Muốn đổi sự kiện hiển thị trên bảng thông tin thì sửa từ khoá ở [cấu hình](cấu%20hình.md), không sửa tab này. Dữ liệu gốc do `LAY_SU_KIEN()` bên data warehouse sinh ra.
