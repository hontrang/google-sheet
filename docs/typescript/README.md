# TypeScript best practices

Tài liệu tham khảo tải về ngày 2026-10-31, lưu nguyên văn (chỉ đổi định dạng) kèm nguồn và giấy phép ở đầu mỗi file.

| File | Nguồn | Nội dung |
| --- | --- | --- |
| [google-ts-style-guide.md](google-ts-style-guide.md) | [Google TypeScript Style Guide](https://google.github.io/styleguide/tsguide.html) (CC BY 3.0) | Style guide đầy đủ: cấu trúc file, import/export, tính năng ngôn ngữ, đặt tên, hệ thống kiểu, comment/JSDoc |
| [ts-handbook-dos-and-donts.md](ts-handbook-dos-and-donts.md) | [TypeScript Handbook — Do's and Don'ts](https://www.typescriptlang.org/docs/handbook/declaration-files/do-s-and-don-ts.html) (CC BY 4.0) | Khuyến nghị chính thức của đội TypeScript về kiểu: `string` thay vì `String`, tránh `any`, `void` cho callback, thứ tự overload, union thay overload |

## Đối chiếu với repo này

Luật thực tế của repo là ESLint (`eslint.config.js`: typescript-eslint strict + stylistic + type-checked) và Prettier — khi tài liệu trên mâu thuẫn với linter, linter thắng. Một số điểm khác biệt đáng chú ý:

- **Prefix `I` cho interface**: Google khuyên không dùng (`IMyInterface`), nhưng repo đặt tên kiểu response là `IResponse*`, `IHttp`, `ISheetSpread` (`src/types/generic.ts`). Giữ nhất quán với code hiện có.
- **Tên định danh tiếng Việt**: Google yêu cầu tên chỉ dùng ký tự ASCII — repo đã viết tiếng Việt không dấu (`layDuLieuTrongO`, `ghiLaiLoVaoSheetVon`), nên vẫn hợp lệ.
- **Module/import**: các phần về ES module, `export`, namespace import không áp dụng nguyên xi cho code chạy trên GAS — build và ts2gas comment hết `import`, mọi file dùng chung một global scope (xem [CLAUDE.md](../../CLAUDE.md) và [architecture.md](../../architecture.md) §2.4).
