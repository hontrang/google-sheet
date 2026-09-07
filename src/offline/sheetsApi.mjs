/**
 * Utility đọc/ghi Google Sheet qua Sheets API v4 từ Node.
 *
 * Xác thực: chạy `login` một lần, token lưu vào `.sheets-token.json` (đã gitignore,
 * chmod 600). Nếu chưa có file đó thì fallback sang refresh token mà `clasp login`
 * lưu trong `.clasprc.json` — nhưng token đó thường đã hết hạn (`invalid_grant`).
 * OAuth client lấy từ `creds.json` cạnh `.clasprc.json` (folder ở biến `mainClasp`).
 *
 * Khác với bản export `.xlsx`, API đọc được:
 *   - công thức sống          → `valueRenderOption=FORMULA`
 *   - pivot table             → `spreadsheets.get` trả về field `pivotTable`
 *   - giá trị lỗi có phân loại → `effectiveValue.errorValue.type`
 *
 * Không nằm trong build pipeline: `npm run build:*` chỉ copy `src/assets/lib`,
 * `src/utility` và folder project, nên file này không bao giờ được đẩy lên GAS.
 *
 * ---- Dùng như CLI ----
 *   node src/offline/sheetsApi.mjs tabs [--id main|dw|<spreadsheetId>]
 *   node src/offline/sheetsApi.mjs values "'bảng thông tin'!G34:J68"
 *   node src/offline/sheetsApi.mjs formulas "'tham chiếu'!V1:V10"
 *   node src/offline/sheetsApi.mjs grid "'bảng thông tin'!G34:J68"
 *   node src/offline/sheetsApi.mjs pivots "bảng thông tin"
 *   node src/offline/sheetsApi.mjs errors "'tham chiếu'!A3:V403"
 *
 * ---- Dùng như library ----
 *   import { layGiaTri, layCongThuc, layPivot } from './src/offline/sheetsApi.mjs';
 *   const rows = await layCongThuc("'bảng thông tin'!G34:J68");
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const GOC_REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const API = 'https://sheets.googleapis.com/v4/spreadsheets';

/** ID 2 sheet đang dùng (xem architecture.md §7). Ghi đè được bằng env. */
export const SPREADSHEET = {
  main: process.env.sheetMainId ?? '1GyxkNiyXantim6R6otooAGd6SLoR4J9Db8nCCx216og',
  dw: process.env.sheetDwId ?? '1yttudZqXZD9URweOzEtkRDda-iUiQEh4a2uNmPlKRfw'
};

/** Đọc `.env` ở gốc repo. Không dùng dotenv để file này chạy được với node trần. */
function docEnv() {
  const duongDan = path.join(GOC_REPO, '.env');
  if (!fs.existsSync(duongDan)) return {};
  const ket = {};
  for (const dong of fs.readFileSync(duongDan, 'utf8').split('\n')) {
    const m = /^\s*([\w.-]+)\s*=\s*(.*)\s*$/.exec(dong);
    if (!m) continue;
    ket[m[1]] = m[2].trim().replace(/^['"]|['"]$/g, '');
  }
  return ket;
}

/** Tìm `.clasprc.json` — ưu tiên biến env `clasprc`, sau đó `mainClasp` trong `.env`. */
export function timClasprc() {
  const env = docEnv();
  const ungVien = [
    process.env.clasprc,
    env.mainClasp && path.join(env.mainClasp, '.clasprc.json'),
    env.dwClasp && path.join(env.dwClasp, '.clasprc.json'),
    path.join(process.env.HOME ?? '', '.clasprc.json')
  ].filter(Boolean);
  const thay = ungVien.find((p) => fs.existsSync(p));
  if (!thay) {
    throw new Error(`Không tìm thấy .clasprc.json. Đã thử:\n  ${ungVien.join('\n  ')}\nChạy \`clasp login\` hoặc set env \`clasprc=<đường dẫn>\`.`);
  }
  return thay;
}

let tokenCache = null;

/** Token do `login` sinh ra, ưu tiên hơn `.clasprc.json`. Đã nằm trong .gitignore. */
const FILE_TOKEN = path.join(GOC_REPO, '.sheets-token.json');

/** OAuth client (Desktop app) — lấy từ `creds.json` cạnh `.clasprc.json`. */
export function docOauthClient() {
  const env = docEnv();
  const ungVien = [process.env.gcpCreds, env.mainClasp && path.join(env.mainClasp, 'creds.json'), path.join(GOC_REPO, 'creds.json')].filter(Boolean);
  const thay = ungVien.find((p) => fs.existsSync(p));
  if (thay) {
    const j = JSON.parse(fs.readFileSync(thay, 'utf8'));
    const c = j.installed ?? j.web ?? j;
    if (c.client_id) return { clientId: c.client_id, clientSecret: c.client_secret, nguon: thay };
  }
  // fallback: client dùng lại từ .clasprc.json
  const rc = JSON.parse(fs.readFileSync(timClasprc(), 'utf8'));
  const s = rc.oauth2ClientSettings ?? {};
  if (!s.clientId) throw new Error(`Không tìm thấy OAuth client. Đã thử:\n  ${ungVien.join('\n  ')}`);
  return { clientId: s.clientId, clientSecret: s.clientSecret, nguon: timClasprc() };
}

/**
 * Đăng nhập lại bằng OAuth loopback flow (dành cho Desktop app client).
 * In ra URL để mở trên trình duyệt, dựng server localhost bắt code, lưu refresh
 * token vào `.sheets-token.json`. Không ghi đè `.clasprc.json` của clasp.
 */
export async function dangNhap({ ghiDuoc = false, cong = 4571, choDoiMs = 300_000 } = {}) {
  const http = await import('node:http');
  const { clientId, clientSecret, nguon } = docOauthClient();
  const redirectUri = `http://localhost:${cong}`;
  const scope = ghiDuoc ? 'https://www.googleapis.com/auth/spreadsheets' : 'https://www.googleapis.com/auth/spreadsheets.readonly';
  const url =
    'https://accounts.google.com/o/oauth2/v2/auth?' +
    new URLSearchParams({ client_id: clientId, redirect_uri: redirectUri, response_type: 'code', scope, access_type: 'offline', prompt: 'consent' });

  console.log(`OAuth client đọc từ: ${nguon}`);
  console.log(`Scope xin cấp     : ${scope}`);
  console.log(`\nMở URL này trên trình duyệt (tài khoản trang11392@gmail.com), bấm Cho phép:\n\n${url}\n`);
  console.log(`Đang chờ redirect về ${redirectUri} … (tối đa ${choDoiMs / 1000}s)`);

  const code = await new Promise((resolve, reject) => {
    const srv = http.createServer((req, res) => {
      const q = new URL(req.url, redirectUri).searchParams;
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(`<h2>${q.get('code') ? 'Xong — quay lại terminal.' : 'Lỗi: ' + q.get('error')}</h2>`);
      srv.close();
      q.get('code') ? resolve(q.get('code')) : reject(new Error(`OAuth bị từ chối: ${q.get('error')}`));
    });
    srv.listen(cong);
    srv.on('error', reject);
    setTimeout(() => {
      srv.close();
      reject(new Error('Hết thời gian chờ đăng nhập.'));
    }, choDoiMs).unref();
  });

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ code, client_id: clientId, client_secret: clientSecret, redirect_uri: redirectUri, grant_type: 'authorization_code' })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Đổi code lấy token thất bại (HTTP ${res.status}): ${data.error_description ?? data.error}`);
  fs.writeFileSync(FILE_TOKEN, JSON.stringify({ refresh_token: data.refresh_token, scope: data.scope, clientId, clientSecret }, null, 2), { mode: 0o600 });
  tokenCache = { token: data.access_token, hetHan: Date.now() + (data.expires_in ?? 3600) * 1000 };
  console.log(`\nĐã lưu refresh token vào ${FILE_TOKEN} (chmod 600, đã có trong .gitignore).`);
  return data.access_token;
}

/** Đổi refresh token lấy access token. Cache trong RAM theo vòng đời tiến trình. */
export async function layAccessToken() {
  if (tokenCache && tokenCache.hetHan > Date.now() + 60_000) return tokenCache.token;

  let refreshToken, clientId, clientSecret, nguon;
  if (fs.existsSync(FILE_TOKEN)) {
    ({ refresh_token: refreshToken, clientId, clientSecret } = JSON.parse(fs.readFileSync(FILE_TOKEN, 'utf8')));
    nguon = FILE_TOKEN;
  } else {
    const creds = JSON.parse(fs.readFileSync(timClasprc(), 'utf8'));
    refreshToken = creds.token?.refresh_token;
    ({ clientId, clientSecret } = creds.oauth2ClientSettings ?? {});
    nguon = timClasprc();
  }
  if (!refreshToken || !clientId) throw new Error(`${nguon} thiếu refresh_token hoặc client. Chạy: node src/offline/sheetsApi.mjs login`);

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken, grant_type: 'refresh_token' })
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(`Refresh token thất bại (HTTP ${res.status}): ${data.error ?? ''} ${data.error_description ?? ''}\nCó thể token đã bị thu hồi — chạy lại \`clasp login\`.`);
  }
  if (!/auth\/spreadsheets/.test(data.scope ?? '')) {
    throw new Error(`Token không có scope spreadsheets. Scope hiện có:\n  ${(data.scope ?? '(rỗng)').split(' ').join('\n  ')}\nChạy: node src/offline/sheetsApi.mjs login`);
  }
  tokenCache = { token: data.access_token, hetHan: Date.now() + (data.expires_in ?? 3600) * 1000 };
  return tokenCache.token;
}

function giaiMaId(id) {
  return SPREADSHEET[id] ?? id ?? SPREADSHEET.main;
}

async function goi(duongDan, { id, ...qs } = {}) {
  const token = await layAccessToken();
  const url = new URL(`${API}/${giaiMaId(id)}${duongDan}`);
  for (const [k, v] of Object.entries(qs)) {
    if (v === undefined) continue;
    for (const x of Array.isArray(v) ? v : [v]) url.searchParams.append(k, x);
  }
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const data = await res.json();
  if (!res.ok) throw new Error(`Sheets API lỗi ${res.status}: ${data.error?.message ?? JSON.stringify(data).slice(0, 300)}`);
  return data;
}

/** Danh sách tab: tên, gid, kích thước, số pivot table. */
export async function layDanhSachTab(id) {
  const data = await goi('', { id, fields: 'properties.title,sheets(properties(sheetId,title,index,hidden,gridProperties))' });
  return {
    tenFile: data.properties?.title,
    tabs: data.sheets.map((s) => ({
      ten: s.properties.title,
      gid: s.properties.sheetId,
      an: s.properties.hidden ?? false,
      dong: s.properties.gridProperties?.rowCount,
      cot: s.properties.gridProperties?.columnCount
    }))
  };
}

/** Giá trị hiển thị của một vùng. */
export async function layGiaTri(range, id) {
  const data = await goi(`/values/${encodeURIComponent(range)}`, { id, valueRenderOption: 'FORMATTED_VALUE' });
  return data.values ?? [];
}

/** Công thức nguyên văn của một vùng (ô không có công thức trả về giá trị thô). */
export async function layCongThuc(range, id) {
  const data = await goi(`/values/${encodeURIComponent(range)}`, { id, valueRenderOption: 'FORMULA' });
  return data.values ?? [];
}

/** Nhiều vùng trong 1 request. */
export async function layNhieuVung(ranges, { id, render = 'FORMULA' } = {}) {
  const data = await goi('/values:batchGet', { id, ranges, valueRenderOption: render });
  return Object.fromEntries(data.valueRanges.map((v) => [v.range, v.values ?? []]));
}

/**
 * Dữ liệu ô đầy đủ: công thức nhập vào, giá trị hiệu dụng, loại lỗi.
 * Trả về mảng 2 chiều các object { a1, congThuc, giaTri, loaiLoi, thongBaoLoi }.
 */
export async function layChiTietO(range, id) {
  const data = await goi('', {
    id,
    ranges: range,
    includeGridData: true,
    fields: 'sheets(properties.title,data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,formattedValue))))'
  });
  const khoi = data.sheets?.[0]?.data?.[0];
  if (!khoi) return [];
  const r0 = khoi.startRow ?? 0;
  const c0 = khoi.startColumn ?? 0;
  return (khoi.rowData ?? []).map((row, i) =>
    (row.values ?? []).map((o, j) => {
      const ev = o.effectiveValue ?? {};
      return {
        a1: `${tenCot(c0 + j)}${r0 + i + 1}`,
        congThuc: o.userEnteredValue?.formulaValue ?? null,
        giaTri: o.formattedValue ?? null,
        loaiLoi: ev.errorValue?.type ?? null,
        thongBaoLoi: ev.errorValue?.message ?? null
      };
    })
  );
}

/** Pivot table trong một tab: vùng neo, vùng nguồn, hàng/cột/giá trị. */
export async function layPivot(tenTab, id) {
  const data = await goi('', {
    id,
    ranges: tenTab,
    includeGridData: true,
    fields: 'sheets(properties.title,data(startRow,startColumn,rowData(values(pivotTable))))'
  });
  const khoi = data.sheets?.[0]?.data?.[0];
  const ket = [];
  (khoi?.rowData ?? []).forEach((row, i) =>
    (row.values ?? []).forEach((o, j) => {
      if (!o.pivotTable) return;
      const p = o.pivotTable;
      ket.push({
        neoTai: `${tenCot((khoi.startColumn ?? 0) + j)}${(khoi.startRow ?? 0) + i + 1}`,
        nguon: p.source,
        hang: p.rows,
        cot: p.columns,
        giaTri: p.values,
        loc: p.criteria ?? p.filterSpecs
      });
    })
  );
  return ket;
}

/** Quét mọi ô lỗi trong một vùng. */
export async function quetLoi(range, id) {
  const chiTiet = await layChiTietO(range, id);
  return chiTiet.flat().filter((o) => o.loaiLoi);
}

/** Ghi giá trị. Cần truyền `xacNhan: true` để tránh ghi nhầm. */
export async function ghiGiaTri(range, values, { id, xacNhan = false, raw = false } = {}) {
  if (!xacNhan) throw new Error('ghiGiaTri() cần { xacNhan: true } — đây là thao tác ghi đè lên sheet thật.');
  const token = await layAccessToken();
  const url = new URL(`${API}/${giaiMaId(id)}/values/${encodeURIComponent(range)}`);
  url.searchParams.set('valueInputOption', raw ? 'RAW' : 'USER_ENTERED');
  const res = await fetch(url, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Sheets API lỗi ${res.status}: ${data.error?.message}`);
  return data;
}

function tenCot(n) {
  let s = '';
  for (let i = n; i >= 0; i = Math.floor(i / 26) - 1) s = String.fromCharCode(65 + (i % 26)) + s;
  return s;
}

/* ------------------------------- CLI ------------------------------- */

async function main() {
  const [lenh, ...rest] = process.argv.slice(2);
  const viTriId = rest.indexOf('--id');
  const id = viTriId >= 0 ? rest[viTriId + 1] : undefined;
  const boQua = viTriId >= 0 ? [viTriId, viTriId + 1] : [];
  const arg = rest.filter((_, i) => !boQua.includes(i)).find((a) => !a.startsWith('--'));
  if (lenh && !['login', 'tabs'].includes(lenh) && !arg) throw new Error(`Lệnh \`${lenh}\` cần một tham số (vùng A1 hoặc tên tab).`);

  switch (lenh) {
    case 'login': {
      await dangNhap({ ghiDuoc: rest.includes('--rw') });
      break;
    }
    case 'tabs': {
      const { tenFile, tabs } = await layDanhSachTab(id);
      console.log(`File: ${tenFile}`);
      for (const t of tabs) console.log(`  ${t.an ? '(ẩn) ' : '     '}${t.ten.padEnd(28)} gid=${String(t.gid).padEnd(12)} ${t.dong}x${t.cot}`);
      break;
    }
    case 'values':
    case 'formulas': {
      const rows = lenh === 'values' ? await layGiaTri(arg, id) : await layCongThuc(arg, id);
      rows.forEach((r, i) => console.log(String(i).padStart(3), JSON.stringify(r)));
      break;
    }
    case 'grid': {
      for (const row of await layChiTietO(arg, id)) {
        for (const o of row) {
          if (!o.congThuc && !o.giaTri && !o.loaiLoi) continue;
          console.log(`${o.a1.padEnd(7)} ${o.loaiLoi ? `[${o.loaiLoi}] ` : ''}${o.congThuc ?? JSON.stringify(o.giaTri)}`);
        }
      }
      break;
    }
    case 'pivots': {
      const ps = await layPivot(arg, id);
      if (!ps.length) console.log(`Tab "${arg}": KHÔNG có pivot table nào.`);
      else console.log(JSON.stringify(ps, null, 2));
      break;
    }
    case 'errors': {
      const loi = await quetLoi(arg, id);
      if (!loi.length) console.log(`${arg}: không có ô lỗi.`);
      for (const o of loi) console.log(`${o.a1.padEnd(8)} ${o.loaiLoi.padEnd(12)} ${o.giaTri ?? ''}  ${o.thongBaoLoi ?? ''}`);
      break;
    }
    default:
      console.log(
        fs
          .readFileSync(fileURLToPath(import.meta.url), 'utf8')
          .split('*/')[0]
          .replace(/^\/\*\*|^ \*ic?/gm, '')
      );
      process.exitCode = 1;
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  main().catch((e) => {
    console.error('LỖI:', e.message);
    process.exitCode = 1;
  });
}
