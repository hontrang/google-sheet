// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { SheetHelper } from './SheetHelper'; // Đảm bảo đường dẫn đúng

export class ZChartHelper {
  public static readonly CHART_ID = 911649750;

  public static updateChart(): void {
    const sheetHelper = new SheetHelper();
    const label = `${sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CHI_TIET_MA, 'F1')} - ${sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CHI_TIET_MA, 'G1')}`;
    const tenMa = `${sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CHI_TIET_MA, 'F1')}`;
    const HIGH_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B5');
    const LOW_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B4');
    const ABS_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'B3');
    const HIGH_VNI: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'C5');
    const LOW_VNI: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'C4');
    const ABS_VNI: number = +sheetHelper.layDuLieuTrongO(SheetHelper.SheetName.SHEET_CAU_HINH, 'C3');
    const chart = this.getChartById(this.CHART_ID, SheetHelper.SheetName.SHEET_CHI_TIET_MA);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SheetName.SHEET_CHI_TIET_MA);

    if (!sheet || !chart) {
      console.error(`Sheet hoặc biểu đồ không tồn tại.`);
      return;
    }

    const updatedChart = chart
      .modify()
      .setOption('title', label)
      .setOption('vAxis.minValue', LOW_MA - ABS_MA * 2)
      .setOption('vAxis.maxValue', HIGH_MA + ABS_MA * 2)
      .setOption('series', { 0: { labelInLegend: tenMa }, 1: { labelInLegend: 'VN-INDEX' } })
      .setOption('vAxes', {
        0: { viewWindow: { min: LOW_MA - ABS_MA * 2, max: HIGH_MA + ABS_MA * 2 } },
        1: { viewWindow: { min: LOW_VNI - ABS_VNI * 2, max: HIGH_VNI + ABS_VNI * 2 } }
      })
      .build();

    sheet.updateChart(updatedChart);
  }

  public static createChart(): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SheetName.SHEET_CHI_TIET_MA);
    if (!sheet) {
      console.error('Không tìm thấy sheet.');
      return;
    }

    const range = sheet.getRange('A1:B8');
    const chartBuilder = sheet.newChart().addRange(range).setChartType(Charts.ChartType.LINE).setPosition(10, 10, 0, 0).setOption('title', 'My Line Chart!');

    const chart = chartBuilder.build();
    sheet.insertChart(chart);
    console.log(chart.getChartId());
  }

  public static getChartById(chartId: number, sheetName: string) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.error(`Sheet ${sheetName} không tồn tại.`);
      return null;
    }

    const charts = sheet.getCharts();
    for (const chart of charts) {
      if (chart.getChartId() === chartId) {
        return chart;
      }
    }

    console.log(`Không tìm thấy biểu đồ có ID ${chartId} trong sheet ${sheetName}.`);
    return null;
  }

  public static removeChartByID(): void {
    const chart = this.getChartById(this.CHART_ID, SheetHelper.SheetName.SHEET_CHI_TIET_MA);
    if (!chart) {
      console.error('Biểu đồ không tồn tại.');
      return;
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SheetName.SHEET_CHI_TIET_MA);
    if (!sheet) {
      console.error(`Sheet ${SheetHelper.SheetName.SHEET_CHI_TIET_MA} không tồn tại.`);
      return;
    }

    sheet.removeChart(chart);
  }

  public static inRaThongTinChart(): void {
    const spreadsheetId: string = SpreadsheetApp.getActiveSpreadsheet().getId();

    const chartsInfo = Sheets.Spreadsheets?.get(spreadsheetId, {
      ranges: [],
      includeGridData: false,
      fields: 'sheets(charts,properties)'
    });

    for (const sheet of chartsInfo?.sheets || []) {
      if (sheet.properties?.title === SheetHelper.SheetName.SHEET_CHI_TIET_MA) {
        if (sheet.charts && sheet.charts.length > 0) {
          for (const chart of sheet.charts) {
            Logger.log(JSON.stringify(chart));
          }
        } else {
          Logger.log(`Không tìm thấy biểu đồ nào trên bảng ${SheetHelper.SheetName.SHEET_CHI_TIET_MA}.`);
        }
        break;
      }
    }
  }
}
