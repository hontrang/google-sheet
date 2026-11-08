/**
 * Utility thao tác Google Drive qua Drive API v3 từ Node.
 *
 * Xác thực dùng chung với sheetsApi.mjs (cùng OAuth client `creds.json`, cùng file
 * `.sheets-token.json`). Token cũ chỉ có scope spreadsheets nên phải `login` lại một
 * lần — lệnh `login` ở đây xin cả scope drive LẪN spreadsheets, nên sau đó sheetsApi.mjs
 * vẫn chạy bình thường với token mới.
 *
 * Thao tác sửa/xoá/chia sẻ file có sẵn đều cần `{ xacNhan: true }`. Tạo mới (upload,
 * tạo thư mục, sao chép) thì không, vì không đụng tới dữ liệu đang có.
 *
 * Không nằm trong build pipeline: `npm run build:*` chỉ copy `src/assets/lib`,
 * `src/utility` và folder project, nên file này không bao giờ được đẩy lên GAS.
 *
 * ---- Dùng như CLI ----
 *   node src/offline/driveApi.mjs login [--readonly]
 *   node src/offline/driveApi.mjs ls [<folderId>]                 # mặc định: My Drive gốc
 *   node src/offline/driveApi.mjs search "name contains 'bctc'"
 *   node src/offline/driveApi.mjs info <fileId>
 *   node src/offline/driveApi.mjs perms <fileId>
 *   node src/offline/driveApi.mjs download <fileId> [<đường dẫn ra>]
 *   node src/offline/driveApi.mjs export <fileId> xlsx|pdf|csv|docx [<đường dẫn ra>]
 *   node src/offline/driveApi.mjs upload <file local> [--parent <folderId>]
 *   node src/offline/driveApi.mjs mkdir <tên> [--parent <folderId>]
 *
 * ---- Dùng như library ----
 *   import { timFile, taiLen, xuatFile } from './src/offline/driveApi.mjs';
 *   const files = await timFile("name contains 'vốn' and trashed = false");
 *   await xuatFile(SPREADSHEET.dw, MIME.xlsx, 'dw.xlsx');
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { dangNhap, layAccessToken, SPREADSHEET } from './sheetsApi.mjs';

export { SPREADSHEET };

const API = 'https://www.googleapis.com/drive/v3';
const API_UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
const LENH_LOGIN = 'node src/offline/driveApi.mjs login';
const TRUONG_MAC_DINH = 'id,name,mimeType,size,modifiedTime,parents,trashed,webViewLink';

export const MIME = {
  thuMuc: 'application/vnd.google-apps.folder',
  googleSheet: 'application/vnd.google-apps.spreadsheet',
  googleDoc: 'application/vnd.google-apps.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pdf: 'application/pdf',
  csv: 'text/csv',
  json: 'application/json',
  txt: 'text/plain',
  png: 'image/png',
  jpg: 'image/jpeg'
};

/** Đăng nhập lại, xin scope drive + spreadsheets trên cùng 1 token. */
export function dangNhapDrive({ chiDoc = false } = {}) {
  const scopes = chiDoc
    ? ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly']
    : ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'];
  return dangNhap({ scopes });
}

function canXacNhan(tenHam, xacNhan, moTa) {
  if (!xacNhan) throw new Error(`${tenHam}() cần { xacNhan: true } — ${moTa}`);
}

/**
 * Gọi Drive API. `qs` là query string (mảng → lặp key), `body` object → JSON.
 * `traVe`: 'json' (mặc định), 'buffer' (tải nội dung file), 'rong' (204 No Content).
 */
async function goi(duongDan, { method = 'GET', qs = {}, body, headers = {}, goc = API, traVe = 'json' } = {}) {
  const token = await layAccessToken({ canScope: /auth\/drive/, lenhLogin: LENH_LOGIN });
  const url = new URL(`${goc}${duongDan}`);
  // luôn bật để đọc được cả file trong shared drive
  url.searchParams.set('supportsAllDrives', 'true');
  for (const [k, v] of Object.entries(qs)) {
    if (v === undefined) continue;
    for (const x of Array.isArray(v) ? v : [v]) url.searchParams.append(k, x);
  }
  const laJson = body !== undefined && !(body instanceof Buffer);
  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...(laJson ? { 'Content-Type': 'application/json' } : {}), ...headers },
    body: laJson ? JSON.stringify(body) : body
  });
  if (!res.ok) {
    const loi = await res.json().catch(() => ({}));
    throw new Error(`Drive API lỗi ${res.status}: ${loi.error?.message ?? res.statusText}`);
  }
  if (traVe === 'buffer') return Buffer.from(await res.arrayBuffer());
  if (traVe === 'rong' || res.status === 204) return null;
  return res.json();
}

/* ------------------------------- Đọc ------------------------------- */

/**
 * Tìm file theo cú pháp query của Drive, tự lật trang tới khi đủ `gioiHan`.
 * @see https://developers.google.com/drive/api/guides/search-files
 */
export async function timFile(q, { gioiHan = 100, truong = TRUONG_MAC_DINH, sapXep = 'folder,name' } = {}) {
  const ket = [];
  let pageToken;
  do {
    const data = await goi('/files', {
      qs: {
        q,
        pageSize: Math.min(gioiHan - ket.length, 1000),
        pageToken,
        orderBy: sapXep,
        fields: `nextPageToken,files(${truong})`,
        includeItemsFromAllDrives: 'true',
        corpora: 'allDrives'
      }
    });
    ket.push(...data.files);
    pageToken = data.nextPageToken;
  } while (pageToken && ket.length < gioiHan);
  return ket;
}

/** File + thư mục con trực tiếp của một thư mục (bỏ qua thùng rác). */
export function layFileTrongThuMuc(thuMucId = 'root', tuyChon = {}) {
  return timFile(`'${thuMucId}' in parents and trashed = false`, tuyChon);
}

/** Tìm theo tên chính xác, có thể giới hạn trong 1 thư mục. */
export function timTheoTen(ten, { thuMucCha, ...tuyChon } = {}) {
  const dk = [`name = '${ten.replace(/['\\]/g, '\\$&')}'`, 'trashed = false'];
  if (thuMucCha) dk.push(`'${thuMucCha}' in parents`);
  return timFile(dk.join(' and '), tuyChon);
}

/** Metadata một file. */
export function layThongTin(fileId, { truong = `${TRUONG_MAC_DINH},owners(emailAddress),description` } = {}) {
  return goi(`/files/${fileId}`, { qs: { fields: truong } });
}

/** Danh sách quyền truy cập của file. */
export async function layQuyen(fileId) {
  const data = await goi(`/files/${fileId}/permissions`, { qs: { fields: 'permissions(id,type,role,emailAddress,domain)' } });
  return data.permissions;
}

/**
 * Tải nội dung file thường (xlsx, pdf, ảnh…). File Google (Sheet/Doc) không có nội dung
 * nhị phân — dùng xuatFile(). Có `dich` thì ghi ra đĩa, không thì trả về Buffer.
 */
export async function taiXuong(fileId, dich) {
  const { mimeType, name } = await layThongTin(fileId, { truong: 'name,mimeType' });
  if (mimeType.startsWith('application/vnd.google-apps.')) {
    throw new Error(`"${name}" là file Google (${mimeType}) — dùng xuatFile(id, MIME.xlsx | MIME.pdf …) thay vì taiXuong().`);
  }
  const buf = await goi(`/files/${fileId}`, { qs: { alt: 'media' }, traVe: 'buffer' });
  if (!dich) return buf;
  fs.writeFileSync(dich, buf);
  return dich;
}

/** Xuất file Google sang định dạng khác (Sheet → xlsx/pdf/csv, Doc → docx/pdf). Giới hạn 10MB. */
export async function xuatFile(fileId, mimeType, dich) {
  const buf = await goi(`/files/${fileId}/export`, { qs: { mimeType }, traVe: 'buffer' });
  if (!dich) return buf;
  fs.writeFileSync(dich, buf);
  return dich;
}

/* ------------------------------- Tạo mới ------------------------------- */

function doanMime(duongDan) {
  return MIME[path.extname(duongDan).slice(1).toLowerCase()] ?? 'application/octet-stream';
}

/** Ghép body multipart/related: phần 1 metadata JSON, phần 2 nội dung file. */
function dongGoiMultipart(metadata, noiDung, mimeType) {
  const ranhGioi = `drive-${Date.now().toString(36)}`;
  const body = Buffer.concat([
    Buffer.from(`--${ranhGioi}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`),
    Buffer.from(`--${ranhGioi}\r\nContent-Type: ${mimeType}\r\n\r\n`),
    noiDung,
    Buffer.from(`\r\n--${ranhGioi}--`)
  ]);
  return { body, headers: { 'Content-Type': `multipart/related; boundary=${ranhGioi}` } };
}

/**
 * Upload file local lên Drive (tạo file mới, không ghi đè file trùng tên).
 * `chuyenThanh` = MIME.googleSheet / MIME.googleDoc để Drive convert khi upload (vd. xlsx → Google Sheet).
 */
export async function taiLen(duongDanLocal, { ten, thuMucCha, mimeType, chuyenThanh, moTa } = {}) {
  const loai = mimeType ?? doanMime(duongDanLocal);
  const metadata = { name: ten ?? path.basename(duongDanLocal), parents: thuMucCha ? [thuMucCha] : undefined, mimeType: chuyenThanh, description: moTa };
  const { body, headers } = dongGoiMultipart(metadata, fs.readFileSync(duongDanLocal), loai);
  return goi('/files', { method: 'POST', goc: API_UPLOAD, qs: { uploadType: 'multipart', fields: TRUONG_MAC_DINH }, body, headers });
}

/** Tạo thư mục. */
export function taoThuMuc(ten, { thuMucCha } = {}) {
  return goi('/files', {
    method: 'POST',
    qs: { fields: TRUONG_MAC_DINH },
    body: { name: ten, mimeType: MIME.thuMuc, parents: thuMucCha ? [thuMucCha] : undefined }
  });
}

/** Lấy thư mục theo tên trong `thuMucCha`, chưa có thì tạo. */
export async function layHoacTaoThuMuc(ten, { thuMucCha = 'root' } = {}) {
  const coSan = (await timTheoTen(ten, { thuMucCha })).find((f) => f.mimeType === MIME.thuMuc);
  return coSan ?? taoThuMuc(ten, { thuMucCha });
}

/** Sao chép file (vd. nhân bản sheet làm bản backup). */
export function saoChep(fileId, { ten, thuMucCha } = {}) {
  return goi(`/files/${fileId}/copy`, {
    method: 'POST',
    qs: { fields: TRUONG_MAC_DINH },
    body: { name: ten, parents: thuMucCha ? [thuMucCha] : undefined }
  });
}

/* ------------------------------- Sửa / xoá ------------------------------- */

/** Ghi đè nội dung file có sẵn bằng file local (giữ nguyên id, tên, quyền). Cần `xacNhan: true`. */
export function capNhatNoiDung(fileId, duongDanLocal, { mimeType, xacNhan = false } = {}) {
  canXacNhan('capNhatNoiDung', xacNhan, 'đây là thao tác ghi đè nội dung file thật trên Drive.');
  return goi(`/files/${fileId}`, {
    method: 'PATCH',
    goc: API_UPLOAD,
    qs: { uploadType: 'media', fields: TRUONG_MAC_DINH },
    body: fs.readFileSync(duongDanLocal),
    headers: { 'Content-Type': mimeType ?? doanMime(duongDanLocal) }
  });
}

/** Đổi tên file/thư mục. Cần `xacNhan: true`. */
export function doiTen(fileId, tenMoi, { xacNhan = false } = {}) {
  canXacNhan('doiTen', xacNhan, 'đây là thao tác sửa file thật trên Drive.');
  return goi(`/files/${fileId}`, { method: 'PATCH', qs: { fields: TRUONG_MAC_DINH }, body: { name: tenMoi } });
}

/** Chuyển file sang thư mục khác (gỡ khỏi mọi thư mục cha hiện tại). Cần `xacNhan: true`. */
export async function diChuyen(fileId, thuMucMoi, { xacNhan = false } = {}) {
  canXacNhan('diChuyen', xacNhan, 'đây là thao tác di chuyển file thật trên Drive.');
  const { parents = [] } = await layThongTin(fileId, { truong: 'parents' });
  return goi(`/files/${fileId}`, {
    method: 'PATCH',
    qs: { addParents: thuMucMoi, removeParents: parents.join(','), fields: TRUONG_MAC_DINH },
    body: {}
  });
}

/** Bỏ vào thùng rác (khôi phục được trong 30 ngày). Cần `xacNhan: true`. */
export function boVaoThungRac(fileId, { xacNhan = false } = {}) {
  canXacNhan('boVaoThungRac', xacNhan, 'đây là thao tác xoá file thật trên Drive.');
  return goi(`/files/${fileId}`, { method: 'PATCH', qs: { fields: TRUONG_MAC_DINH }, body: { trashed: true } });
}

/** Lấy file ra khỏi thùng rác. */
export function khoiPhuc(fileId) {
  return goi(`/files/${fileId}`, { method: 'PATCH', qs: { fields: TRUONG_MAC_DINH }, body: { trashed: false } });
}

/** Xoá vĩnh viễn, KHÔNG qua thùng rác, không khôi phục được. Cần `xacNhan: true`. */
export async function xoaVinhVien(fileId, { xacNhan = false } = {}) {
  canXacNhan('xoaVinhVien', xacNhan, 'file sẽ mất hẳn, không qua thùng rác.');
  await goi(`/files/${fileId}`, { method: 'DELETE', traVe: 'rong' });
}

/**
 * Cấp quyền truy cập. `loai`: user | group | domain | anyone; `vaiTro`: reader | commenter | writer.
 * Cần `xacNhan: true` vì đây là thao tác chia sẻ dữ liệu ra ngoài.
 */
export function chiaSe(fileId, { email, loai = 'user', vaiTro = 'reader', guiThongBao = false, xacNhan = false } = {}) {
  canXacNhan('chiaSe', xacNhan, 'đây là thao tác chia sẻ file thật ra bên ngoài.');
  return goi(`/files/${fileId}/permissions`, {
    method: 'POST',
    qs: { sendNotificationEmail: String(guiThongBao), fields: 'id,type,role,emailAddress' },
    body: { type: loai, role: vaiTro, emailAddress: email }
  });
}

/** Thu hồi một quyền (id lấy từ layQuyen()). Cần `xacNhan: true`. */
export async function thuHoiQuyen(fileId, quyenId, { xacNhan = false } = {}) {
  canXacNhan('thuHoiQuyen', xacNhan, 'đây là thao tác sửa quyền file thật.');
  await goi(`/files/${fileId}/permissions/${quyenId}`, { method: 'DELETE', traVe: 'rong' });
}

/* ------------------------------- CLI ------------------------------- */

function inDanhSach(files) {
  if (!files.length) console.log('(trống)');
  for (const f of files) {
    const loai = f.mimeType === MIME.thuMuc ? '[dir]' : f.mimeType.replace(/^application\/vnd\.google-apps\./, 'g:').slice(0, 14);
    console.log(`${loai.padEnd(15)} ${f.id.padEnd(45)} ${(f.size ? `${Math.ceil(f.size / 1024)}KB` : '').padStart(8)}  ${f.name}`);
  }
}

async function main() {
  const [lenh, ...rest] = process.argv.slice(2);
  const viTriParent = rest.indexOf('--parent');
  const thuMucCha = viTriParent >= 0 ? rest[viTriParent + 1] : undefined;
  const boQua = viTriParent >= 0 ? [viTriParent, viTriParent + 1] : [];
  const args = rest.filter((a, i) => !boQua.includes(i) && !a.startsWith('--'));
  const canArg = (n) => {
    if (args.length < n) throw new Error(`Lệnh \`${lenh}\` cần ${n} tham số.`);
  };

  switch (lenh) {
    case 'login':
      await dangNhapDrive({ chiDoc: rest.includes('--readonly') });
      break;
    case 'ls':
      inDanhSach(await layFileTrongThuMuc(args[0] ?? 'root'));
      break;
    case 'search':
      canArg(1);
      inDanhSach(await timFile(args[0]));
      break;
    case 'info':
      canArg(1);
      console.log(JSON.stringify(await layThongTin(args[0]), null, 2));
      break;
    case 'perms':
      canArg(1);
      for (const p of await layQuyen(args[0])) console.log(`${p.role.padEnd(10)} ${p.type.padEnd(7)} ${p.emailAddress ?? p.domain ?? ''}  (id=${p.id})`);
      break;
    case 'download': {
      canArg(1);
      const dich = args[1] ?? (await layThongTin(args[0], { truong: 'name' })).name;
      console.log(`Đã lưu: ${await taiXuong(args[0], dich)}`);
      break;
    }
    case 'export': {
      canArg(2);
      const mime = MIME[args[1]] ?? args[1];
      const dich = args[2] ?? `${(await layThongTin(args[0], { truong: 'name' })).name}.${args[1]}`;
      console.log(`Đã lưu: ${await xuatFile(args[0], mime, dich)}`);
      break;
    }
    case 'upload': {
      canArg(1);
      const f = await taiLen(args[0], { thuMucCha });
      console.log(`Đã tải lên: ${f.name}  id=${f.id}\n${f.webViewLink}`);
      break;
    }
    case 'mkdir': {
      canArg(1);
      const f = await taoThuMuc(args[0], { thuMucCha });
      console.log(`Đã tạo thư mục: ${f.name}  id=${f.id}\n${f.webViewLink}`);
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
