/**
 * Chèn mã chứng khoán vào cột cuối dòng 1 của 4 ma trận trong sheet "data warehouse" (projectDw):
 * Giá, Khối Lượng, KN Mua, KN Bán.
 *
 *   node .claude/skills/them-ma-dw/them-ma.mjs <MA> [<MA> ...] [--apply]
 *
 * Mặc định chỉ đọc và in kế hoạch. `--apply` mới ghi lên sheet thật.
 * Tab nào đã có mã thì bỏ qua, nên chạy lại không bị trùng.
 *
 * Các hàm cập nhật hằng ngày (15:50) dò mã theo tên ở dòng 1 của từng tab; layGiaHangNgay()
 * lấy danh sách mã từ Khối Lượng nhưng ghi vào Giá — vì vậy mã phải có ở cả 4 tab.
 */

import { SPREADSHEET, layDanhSachTab, ghiGiaTri, capNhatLo, tenCot } from '../../../src/offline/sheetsApi.mjs';
import { MA_TRAN, ID, COT_IMPORT_CUOI, VI_TRI_IMPORT_CUOI, docMa, traMaVndirect, viTriCuoi, docDongTieuDe } from './chung.mjs';

async function main() {
  const args = process.argv.slice(2);
  const apply = args.includes('--apply');
  const danhSachMa = docMa(args);
  if (danhSachMa.length === 0) throw new Error('Thiếu mã. Dùng: node .claude/skills/them-ma-dw/them-ma.mjs <MA> [<MA> ...] [--apply]');

  console.log(`Sheet data warehouse: ${SPREADSHEET[ID]}`);
  console.log(`Thêm mã: ${danhSachMa.join(', ')}${apply ? '  — CHẾ ĐỘ GHI (--apply)' : '  — xem trước, chưa ghi gì (thêm --apply để ghi)'}\n`);

  for (const ma of danhSachMa) {
    const info = await traMaVndirect(ma);
    if (!info) throw new Error(`vndirect không biết mã ${ma}. Kiểm tra lại mã.`);
    const canhBao = info.type !== 'STOCK' || info.status !== 'listed' ? '  ⚠️ không phải cổ phiếu đang niêm yết — hàm cập nhật có thể không lấy được dữ liệu' : '';
    console.log(`vndirect ${ma}: ${info.companyName} | ${info.floor} | ${info.type} | ${info.status} | niêm yết ${info.listedDate}${canhBao}`);
  }

  const { tabs } = await layDanhSachTab(ID);
  const luoi = Object.fromEntries(tabs.map((t) => [t.ten, t]));
  const tieuDe = await docDongTieuDe();

  const themCot = []; // batchUpdate: mở rộng lưới khi hết cột
  const ghi = []; // [range, values]
  const ngoaiImport = new Set();
  console.log('\nKế hoạch:');
  MA_TRAN.forEach((tab, k) => {
    const hang = tieuDe[k];
    const canThem = danhSachMa.filter((ma) => !hang.includes(ma));
    for (const ma of danhSachMa.filter((m) => hang.includes(m))) console.log(`  ✔ ${tab}: ${ma} đã có ở ${tenCot(hang.indexOf(ma))}1`);
    if (canThem.length === 0) return;
    const cotDau = viTriCuoi(hang) + 1;
    const cotCuoi = cotDau + canThem.length - 1;
    canThem.forEach((ma, i) => cotDau + i > VI_TRI_IMPORT_CUOI && ngoaiImport.add(ma));
    const thieu = cotCuoi + 1 - luoi[tab].cot;
    if (thieu > 0) themCot.push({ appendDimension: { sheetId: luoi[tab].gid, dimension: 'COLUMNS', length: thieu } });
    ghi.push([`'${tab}'!${tenCot(cotDau)}1:${tenCot(cotCuoi)}1`, [canThem]]);
    console.log(
      `  + ${tab}: ghi ${canThem.join(', ')} vào ${tenCot(cotDau)}1${cotCuoi > cotDau ? `:${tenCot(cotCuoi)}1` : ''}${thieu > 0 ? ` (thêm ${thieu} cột, lưới đang có ${luoi[tab].cot} cột)` : ''}`
    );
  });
  if (ngoaiImport.size) {
    console.log(
      `\n⚠️  ${[...ngoaiImport].join(', ')} sẽ nằm sau cột ${COT_IMPORT_CUOI}: file chứng khoán chỉ import tới ${COT_IMPORT_CUOI} (tab Giá/Khối Lượng/KN Mua/KN Bán, ô A1–A2),`
    );
    console.log(`    nên mã này vẫn được cập nhật trong DW nhưng KHÔNG lên dashboard cho tới khi nới vùng import hoặc xoá bớt mã (xoa-ma.mjs).`);
  }

  if (ghi.length === 0) return console.log('\nKhông có gì để ghi — mã đã có ở cả 4 tab.');
  if (!apply) return console.log('\nChưa ghi gì. Chạy lại với --apply để thực hiện.');

  if (themCot.length) await capNhatLo(themCot, { id: ID, xacNhan: true });
  for (const [range, values] of ghi) await ghiGiaTri(range, values, { id: ID, xacNhan: true });
  console.log(`\nĐã ghi ${ghi.length} tab${themCot.length ? `, thêm cột ở ${themCot.length} tab` : ''}. Kiểm tra lại:`);

  const sau = await docDongTieuDe();
  MA_TRAN.forEach((tab, k) => {
    for (const ma of danhSachMa) {
      const i = sau[k].indexOf(ma);
      console.log(`  ${i >= 0 ? '✔' : '✘'} ${tab}: ${ma}${i >= 0 ? ` ở ${tenCot(i)}1` : ' KHÔNG thấy'}`);
    }
  });
}

main().catch((e) => {
  console.error(`Lỗi: ${e.message}`);
  process.exit(1);
});
