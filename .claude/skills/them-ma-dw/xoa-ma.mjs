/**
 * Xoá mã chứng khoán khỏi 4 ma trận trong sheet "data warehouse" (projectDw): xoá CẢ CỘT của mã
 * ở Giá, Khối Lượng, KN Mua, KN Bán — mất toàn bộ lịch sử giá / khối lượng / khối ngoại của mã đó.
 *
 *   node .claude/skills/them-ma-dw/xoa-ma.mjs <MA> [<MA> ...] [--apply] [--xoa-ca-lich-su=<MA>,<MA>]
 *
 * Mặc định chỉ đọc và in kế hoạch. `--apply` mới xoá, và trước khi xoá lưu toàn bộ cột
 * (giá trị + công thức) ra backup-dw/<thời điểm>-<mã>.json ở gốc repo (đã gitignore).
 *
 * Khoá an toàn: mục đích chính là gỡ mã vừa thêm để test (cột chỉ có tiêu đề). Cột nào đã có
 * dữ liệu lịch sử thì bị CHẶN, trừ khi mã đó được ghi rõ trong `--xoa-ca-lich-su=`. Chỉ cần một
 * mã bị chặn là không xoá gì cả.
 *
 * Các mã phía sau dịch sang trái một cột. Không công thức nào bị lệch vì:
 *   - các hàm GAS và công thức ở tab dữ liệu dò mã theo tên (MATCH trên dòng 1);
 *   - file chứng khoán import cả tiêu đề lẫn dữ liệu cùng một vùng A:OR nên dịch cùng nhau;
 *   - tham chiếu trực tiếp (TRUY VAN: 'Giá'!IH:IH, …) được Google Sheets tự dịch —
 *     trừ khi nó trỏ đúng cột bị xoá thì thành #REF! (script cảnh báo trước).
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { SPREADSHEET, layDanhSachTab, layNhieuVung, capNhatLo, tenCot } from '../../../src/offline/sheetsApi.mjs';
import { MA_TRAN, ID, docMa, docDongTieuDe } from './chung.mjs';

const GOC_REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const THU_MUC_BACKUP = path.join(GOC_REPO, 'backup-dw');

/** Tham chiếu trực tiếp tới cột của 4 ma trận trong tab TRUY VAN, dạng { 'Giá': Set(['IH', …]) }. */
async function docThamChieuTruyVan() {
  const ketQua = Object.fromEntries(MA_TRAN.map((t) => [t, new Map()]));
  const vung = Object.values(await layNhieuVung(["'TRUY VAN'!1:2"], { id: ID, render: 'FORMULA' }))[0];
  vung.forEach((hang, i) =>
    hang.forEach((o, j) => {
      for (const m of String(o).matchAll(/'([^']+)'!\$?([A-Z]+)\$?\d*:\$?([A-Z]+)/g)) {
        if (ketQua[m[1]]) ketQua[m[1]].set(m[2], `TRUY VAN!${tenCot(j)}${i + 1}`);
      }
    })
  );
  return ketQua;
}

async function main() {
  const args = process.argv.slice(2);
  const apply = args.includes('--apply');
  const danhSachMa = docMa(args);
  const choPhepMatLichSu = new Set(
    args
      .filter((a) => a.startsWith('--xoa-ca-lich-su='))
      .flatMap((a) => a.slice('--xoa-ca-lich-su='.length).split(','))
      .map((m) => m.trim().toUpperCase())
      .filter(Boolean)
  );
  if (danhSachMa.length === 0) throw new Error('Thiếu mã. Dùng: node .claude/skills/them-ma-dw/xoa-ma.mjs <MA> [<MA> ...] [--apply]');

  console.log(`Sheet data warehouse: ${SPREADSHEET[ID]}`);
  console.log(`XOÁ mã: ${danhSachMa.join(', ')}${apply ? '  — CHẾ ĐỘ XOÁ (--apply)' : '  — xem trước, chưa xoá gì (thêm --apply để xoá)'}\n`);

  const { tabs } = await layDanhSachTab(ID);
  const luoi = Object.fromEntries(tabs.map((t) => [t.ten, t]));
  const tieuDe = await docDongTieuDe();
  const truyVan = await docThamChieuTruyVan();
  const cotDuLieu = Object.values(await layNhieuVung(["'dữ liệu'!A:A"], { id: ID, render: 'FORMATTED_VALUE' }))[0].map((r) => r[0]);

  // Kế hoạch: vị trí cột của từng mã ở từng tab
  const viTri = []; // { tab, ma, cot }
  const canhBao = [];
  console.log('Kế hoạch:');
  MA_TRAN.forEach((tab, k) => {
    for (const ma of danhSachMa) {
      const cot = tieuDe[k].indexOf(ma);
      if (cot < 0) {
        console.log(`  · ${tab}: không có ${ma} — bỏ qua`);
        continue;
      }
      if (cot === 0) throw new Error(`${tab}: ${ma} nằm ở cột A (cột Ngày) — dừng lại, không xoá.`);
      viTri.push({ tab, ma, cot });
      console.log(`  − ${tab}: xoá cột ${tenCot(cot)} (${ma})`);
      const oTro = truyVan[tab].get(tenCot(cot));
      if (oTro) canhBao.push(`${oTro} đang trỏ thẳng vào '${tab}'!${tenCot(cot)} (${ma}) — sẽ thành #REF!`);
    }
  });
  for (const ma of danhSachMa) {
    const dong = cotDuLieu.indexOf(ma);
    if (dong >= 0) canhBao.push(`dữ liệu!A${dong + 1} vẫn có ${ma} — các cột giá/khối lượng của dòng đó sẽ ra #N/A (skill không xoá dòng này)`);
  }
  canhBao.push('File chứng khoán: nếu tab dữ liệu của file đó còn mã này, các cột tra theo mã sẽ ra #N/A.');

  if (viTri.length === 0) return console.log('\nKhông có gì để xoá — các mã không có ở tab nào.');

  // Đếm lịch sử sẽ mất
  const cotDayDu = await layNhieuVung(
    viTri.map(({ tab, cot }) => `'${tab}'!${tenCot(cot)}:${tenCot(cot)}`),
    { id: ID, render: 'FORMULA' }
  );
  const duLieuCot = Object.values(cotDayDu);
  viTri.forEach((v, i) => (v.giaTri = duLieuCot[i]));
  console.log('\nDữ liệu sẽ mất:');
  for (const v of viTri) {
    v.soO = v.giaTri.slice(1).filter((r) => String(r[0] ?? '') !== '').length;
    console.log(`  ${v.tab} ${v.ma}: ${v.soO === 0 ? 'không có lịch sử (chỉ tiêu đề)' : `${v.soO} ô có dữ liệu`}`);
  }

  // Khoá an toàn: mã có lịch sử phải được cho phép rõ ràng
  const biChan = [...new Set(viTri.filter((v) => v.soO > 0 && !choPhepMatLichSu.has(v.ma)).map((v) => v.ma))];

  console.log('\nCảnh báo:');
  for (const c of canhBao) console.log(`  ⚠️ ${c}`);

  if (biChan.length) {
    console.log(`\n🔒 BỊ CHẶN: ${biChan.join(', ')} đã có dữ liệu lịch sử. Script chỉ xoá cột chưa có lịch sử (mã vừa thêm để test).`);
    console.log(`   Chỉ khi người dùng xác nhận rõ muốn mất lịch sử của đúng mã đó mới thêm: --xoa-ca-lich-su=${biChan.join(',')}`);
    if (apply) throw new Error('Không xoá gì — còn mã bị chặn.');
    return;
  }

  if (!apply) return console.log('\nChưa xoá gì. Chạy lại với --apply để xoá (sẽ lưu backup trước).');

  // Backup trước khi xoá
  fs.mkdirSync(THU_MUC_BACKUP, { recursive: true });
  const thoiDiem = new Date().toISOString().replace(/[:.]/g, '-');
  const fileBackup = path.join(THU_MUC_BACKUP, `${thoiDiem}-${danhSachMa.join('-')}.json`);
  fs.writeFileSync(
    fileBackup,
    JSON.stringify({ spreadsheetId: SPREADSHEET[ID], thoiDiem, cot: viTri.map(({ tab, ma, cot, giaTri }) => ({ tab, ma, cot: tenCot(cot), giaTri })) }, null, 1)
  );
  console.log(`\nĐã lưu backup: ${fileBackup}`);

  // Đọc lại tiêu đề ngay trước khi xoá để chắc vị trí chưa đổi
  const truocXoa = await docDongTieuDe();
  for (const v of viTri) {
    if (truocXoa[MA_TRAN.indexOf(v.tab)][v.cot] !== v.ma) throw new Error(`${v.tab}: cột ${tenCot(v.cot)} không còn là ${v.ma} — sheet vừa thay đổi, dừng lại (chưa xoá gì).`);
  }

  // Trong mỗi tab xoá từ phải sang trái để index của các cột còn lại không đổi
  const requests = [...viTri]
    .sort((a, b) => a.tab.localeCompare(b.tab) || b.cot - a.cot)
    .map(({ tab, cot }) => ({ deleteDimension: { range: { sheetId: luoi[tab].gid, dimension: 'COLUMNS', startIndex: cot, endIndex: cot + 1 } } }));
  await capNhatLo(requests, { id: ID, xacNhan: true });
  console.log(`Đã xoá ${requests.length} cột. Kiểm tra lại:`);

  const sau = await docDongTieuDe();
  MA_TRAN.forEach((tab, k) => {
    for (const ma of danhSachMa) console.log(`  ${sau[k].includes(ma) ? '✘ vẫn còn' : '✔ đã hết'} ${ma} ở ${tab}`);
  });
}

main().catch((e) => {
  console.error(`Lỗi: ${e.message}`);
  process.exit(1);
});
