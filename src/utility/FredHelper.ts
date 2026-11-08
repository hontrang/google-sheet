import { IFredObservation, ILaiSuatMucTieuFed, IResponseFredObservations } from '@src/types/generic';

/**
 * Gọi FRED API (Federal Reserve Bank of St. Louis) — https://fred.stlouisfed.org/docs/api/fred/
 * Cần API key miễn phí, lưu ở Script Properties `FRED_API_KEY` (không viết vào code):
 * GAS editor → Project Settings → Script properties, hoặc chạy `FredHelper.luuApiKey('<key>')` một lần.
 */
export class FredHelper {
  public static readonly tenThuocTinhApiKey = 'FRED_API_KEY';
  public static readonly urlQuanSat = 'https://api.stlouisfed.org/fred/series/observations';
  /** Series biên độ lãi suất mục tiêu Fed Funds (%, theo ngày). */
  public static readonly seriesCanTren = 'DFEDTARU';
  public static readonly seriesCanDuoi = 'DFEDTARL';

  private readonly apiKey: string;

  constructor(apiKey = PropertiesService.getScriptProperties().getProperty(FredHelper.tenThuocTinhApiKey) ?? '') {
    this.apiKey = apiKey;
  }

  public static luuApiKey(apiKey: string): void {
    PropertiesService.getScriptProperties().setProperty(FredHelper.tenThuocTinhApiKey, apiKey.trim());
  }

  /**
   * Các quan sát của một series. Bỏ qua ngày không có số liệu (FRED trả `"."`).
   * @param tuy.tuNgay / denNgay `yyyy-MM-dd`; tuy.moiNhatTruoc = true → sắp xếp mới nhất trước; tuy.gioiHan số dòng tối đa.
   */
  layQuanSat(seriesId: string, tuy: { tuNgay?: string; denNgay?: string; moiNhatTruoc?: boolean; gioiHan?: number } = {}): IFredObservation[] {
    if (!this.apiKey) throw new Error(`Chưa có FRED API key. Lưu vào Script Properties "${FredHelper.tenThuocTinhApiKey}".`);
    // Tên tham số do FRED quy định (snake_case), nên để dạng cặp [tên, giá trị] thay vì object literal
    const thamSo: [string, string | undefined][] = [
      ['series_id', seriesId],
      ['api_key', this.apiKey],
      ['file_type', 'json'],
      ['observation_start', tuy.tuNgay],
      ['observation_end', tuy.denNgay],
      ['sort_order', tuy.moiNhatTruoc ? 'desc' : 'asc'],
      ['limit', tuy.gioiHan === undefined ? undefined : String(tuy.gioiHan)]
    ];
    const query = thamSo
      .filter(([, giaTri]) => giaTri !== undefined)
      .map(([ten, giaTri]) => `${ten}=${encodeURIComponent(giaTri ?? '')}`)
      .join('&');
    const response = UrlFetchApp.fetch(`${FredHelper.urlQuanSat}?${query}`, { muteHttpExceptions: true });
    return this.phanTichQuanSat(response.getContentText(), response.getResponseCode());
  }

  /** Quan sát mới nhất có số liệu của một series. */
  layGiaTriMoiNhat(seriesId: string): IFredObservation | undefined {
    return this.layQuanSat(seriesId, { moiNhatTruoc: true, gioiHan: 10 })[0];
  }

  /** Biên độ lãi suất mục tiêu Fed Funds hiện tại, ví dụ { canDuoi: 3.75, canTren: 4, ngay: '2026-11-06' }. */
  layLaiSuatMucTieuFed(): ILaiSuatMucTieuFed | undefined {
    const canTren = this.layGiaTriMoiNhat(FredHelper.seriesCanTren);
    const canDuoi = this.layGiaTriMoiNhat(FredHelper.seriesCanDuoi);
    if (!canTren || !canDuoi) return undefined;
    return { canDuoi: canDuoi.giaTri, canTren: canTren.giaTri, ngay: canTren.ngay < canDuoi.ngay ? canTren.ngay : canDuoi.ngay };
  }

  /** Tách khỏi `layQuanSat()` để test được với JSON cố định. */
  phanTichQuanSat(noiDung: string, maHttp = 200): IFredObservation[] {
    const data = JSON.parse(noiDung) as IResponseFredObservations;
    if (maHttp !== 200 || data.error_code) throw new Error(`FRED API lỗi ${data.error_code ?? maHttp}: ${data.error_message ?? noiDung.slice(0, 200)}`);
    return (data.observations ?? []).filter((o) => o.value !== '.' && o.value !== '').map((o) => ({ ngay: o.date, giaTri: Number(o.value) }));
  }
}
