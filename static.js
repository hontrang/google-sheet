//// Khai bao bien toan cuc /////
let SHEET_THAM_CHIEU = "tham chiếu";
let SHEET_BANG_THONG_TIN = "bảng thông tin";
let SHEET_TIN_TUC = "tin tức";
let SHEET_DU_LIEU = "dữ liệu";
let SHEET_CHI_TIET_MA = "chi tiết mã";
let KICH_THUOC_MANG_PHU = 10;
let HEADER_MA = "Tên mã";
let HEADER_KHOI_LUONG = "Khối lượng"
let HEADER_CODE = "Tên sự kiện";
let HEADER_DATE = "Thời gian diễn ra sự kiện";
let HEADER_LINK = "Link sự kiện";
let mangPhu = new Array();
let mang_du_lieu_chinh = new Array();
let response;
let object;
let url;
let range;
let tenMa;


let OPTIONS = {
  method: 'GET',
  headers: {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
  }
}

let OPTIONS_HTML = {
  method: 'GET',
  headers: {
    'Content-Type': 'text/html',
    'Accept': 'text/html',
  }
}