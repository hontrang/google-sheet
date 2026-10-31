/** Phần dùng chung của them-ma.mjs và xoa-ma.mjs. */

import { layNhieuVung } from '../../../src/offline/sheetsApi.mjs';

// Tên tab, khớp SheetHelper.sheetName (src/utility/SheetHelper.ts)
export const MA_TRAN = ['Giá', 'Khối Lượng', 'KN Mua', 'KN Bán'];
export const ID = 'dw';

/** Tên cột (A, Z, AA, OR…) → index 0-based. Ngược với tenCot() của sheetsApi.mjs. */
export const viTriCot = (ten) => [...ten].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

/**
 * File chứng khoán import 4 tab này bằng vùng cố định `A1:OR…` (tab Giá/Khối Lượng/KN Mua/KN Bán
 * của file chứng khoán), nên mã ở cột sau OR không lên dashboard.
 */
export const COT_IMPORT_CUOI = 'OR';
export const VI_TRI_IMPORT_CUOI = viTriCot(COT_IMPORT_CUOI);

export function docMa(args) {
  const danhSachMa = [...new Set(args.filter((a) => !a.startsWith('--')).map((a) => a.trim().toUpperCase()))];
  const sai = danhSachMa.filter((ma) => !/^[A-Z][A-Z0-9]{2}$/.test(ma));
  if (sai.length) throw new Error(`Không giống mã cổ phiếu (3 ký tự, bắt đầu bằng chữ): ${sai.join(', ')}`);
  return danhSachMa;
}

export async function traMaVndirect(ma) {
  const url = `https://api-finfo.vndirect.com.vn/v4/stocks?q=code:${encodeURIComponent(ma)}&fields=code,companyName,floor,type,status,listedDate`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`vndirect trả về HTTP ${res.status}`);
  return (await res.json()).data?.[0];
}

/** Vị trí (0-based) phần tử không rỗng cuối cùng. */
export const viTriCuoi = (mang) => mang.reduce((cuoi, x, i) => (String(x ?? '').trim() !== '' ? i : cuoi), -1);

/** Dòng 1 của 4 ma trận, theo thứ tự MA_TRAN. */
export async function docDongTieuDe() {
  const vung = await layNhieuVung(
    MA_TRAN.map((t) => `'${t}'!1:1`),
    { id: ID, render: 'FORMATTED_VALUE' }
  );
  return Object.values(vung).map((v) => v[0] ?? []);
}
