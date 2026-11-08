import * as Cheerio from 'cheerio';
import * as luxon from 'luxon';
import { IFomcMeeting } from '@src/types/generic';
import { DateHelper } from '@utils/DateHelper';

/**
 * Lịch họp FOMC lấy từ trang chính thức của Fed (HTML, không cần API key).
 * Trang liệt kê năm hiện tại và vài năm trước; mỗi năm một khối `.panel-default`, mỗi cuộc họp một dòng `.fomc-meeting`.
 */
export class FomcHelper {
  public static readonly urlTrangFed = 'https://www.federalreserve.gov';
  public static readonly urlLichFomc = 'https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm';

  private static readonly thang: Record<string, number> = {
    jan: 1,
    feb: 2,
    mar: 3,
    apr: 4,
    may: 5,
    jun: 6,
    jul: 7,
    aug: 8,
    sep: 9,
    oct: 10,
    nov: 11,
    dec: 12
  };

  /** Toàn bộ cuộc họp trên trang, sắp xếp theo ngày kết thúc tăng dần. */
  layLichFomc(): IFomcMeeting[] {
    const response = UrlFetchApp.fetch(FomcHelper.urlLichFomc, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error(`Không tải được lịch FOMC (HTTP ${response.getResponseCode()})`);
    return this.phanTichLichFomc(response.getContentText());
  }

  /**
   * Cuộc họp sắp tới: ngày kết thúc >= hôm nay (cuộc họp đang diễn ra cũng tính).
   * Truyền `lich` (kết quả `layLichFomc()`) để không tải trang lần nữa.
   */
  layCuocHopTiepTheo(homNay = DateHelper.layNgayHienTai('yyyy-MM-dd'), lich = this.layLichFomc()): IFomcMeeting | undefined {
    return lich.find((cuocHop) => cuocHop.ngayKetThuc >= homNay);
  }

  /** Cuộc họp gần nhất đã kết thúc trước hôm nay. Truyền `lich` để không tải trang lần nữa. */
  layCuocHopGanNhat(homNay = DateHelper.layNgayHienTai('yyyy-MM-dd'), lich = this.layLichFomc()): IFomcMeeting | undefined {
    return lich.filter((cuocHop) => cuocHop.ngayKetThuc < homNay).pop();
  }

  /**
   * Thời điểm Fed công bố quyết định (14:00 giờ New York ngày cuối cuộc họp) đổi sang múi giờ khác, mặc định giờ Việt Nam.
   * Ví dụ cuộc họp kết thúc 2026-12-09 → `02:00 10/12/2026` (luxon tự tính giờ mùa hè/mùa đông của Mỹ).
   */
  gioCongBoQuyetDinh(cuocHop: IFomcMeeting, dinhDang = 'HH:mm dd/MM/yyyy', muiGio = 'Asia/Ho_Chi_Minh'): string {
    return luxon.DateTime.fromISO(`${cuocHop.ngayKetThuc}T14:00`, { zone: 'America/New_York' }).setZone(muiGio).toFormat(dinhDang);
  }

  /** Tách khỏi `layLichFomc()` để test được với HTML cố định. */
  phanTichLichFomc(html: string): IFomcMeeting[] {
    const $ = Cheerio.load(html);
    const ketQua: IFomcMeeting[] = [];
    $('.panel-default').each((_, panel) => {
      const nam = Number(/(\d{4}) FOMC Meetings/.exec($(panel).find('.panel-heading').text())?.[1]);
      if (!nam) return;
      $(panel)
        .find('.fomc-meeting')
        .each((_, dong) => {
          const thang = $(dong).find('.fomc-meeting__month').text().trim();
          const ngay = $(dong).find('.fomc-meeting__date').text().trim();
          const cuocHop = this.phanTichNgay(nam, thang, ngay);
          if (!cuocHop) return;
          const link = (mau: RegExp) => {
            const href = $(dong)
              .find('a')
              .toArray()
              .map((a) => $(a).attr('href') ?? '')
              .find((h) => mau.test(h));
            return href ? FomcHelper.urlTrangFed + href : undefined;
          };
          ketQua.push({
            ...cuocHop,
            linkStatement: link(/pressreleases\/monetary\d{8}a\.htm/),
            linkHopBao: link(/fomcpres+conf\d{8}\.htm/),
            linkBienBan: link(/fomcminutes\d{8}\.htm/)
          });
        });
    });
    return ketQua.sort((a, b) => a.ngayKetThuc.localeCompare(b.ngayKetThuc));
  }

  /**
   * Đổi ô tháng + ô ngày của trang Fed sang ngày cụ thể.
   * Ví dụ: `January` + `27-28`, `March` + `17-18*` (có SEP), `Apr/May` + `30-1` (qua tháng), `March` + `9 (unscheduled)`.
   */
  phanTichNgay(nam: number, oThang: string, oNgay: string): Omit<IFomcMeeting, 'linkStatement' | 'linkHopBao' | 'linkBienBan'> | undefined {
    const [thangDau, thangCuoi = thangDau] = oThang.split('/').map((t) => FomcHelper.thang[t.trim().slice(0, 3).toLowerCase()]);
    const ngay = /(\d{1,2})(?:\s*-\s*(\d{1,2}))?/.exec(oNgay);
    if (!thangDau || !thangCuoi || !ngay) return undefined;
    const ngayDau = Number(ngay[1]);
    const ngayCuoi = ngay[2] ? Number(ngay[2]) : ngayDau;
    const dinhDang = (thang: number, ngayTrongThang: number) => `${nam}-${String(thang).padStart(2, '0')}-${String(ngayTrongThang).padStart(2, '0')}`;
    return {
      nam,
      ngayBatDau: dinhDang(thangDau, ngayDau),
      ngayKetThuc: dinhDang(thangCuoi, ngayCuoi),
      coDuBaoKinhTe: oNgay.includes('*'),
      ghiChu: /\(([^)]*)\)/.exec(oNgay)?.[0] ?? ''
    };
  }
}
