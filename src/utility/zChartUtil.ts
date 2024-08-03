// eslint-disable-next-line @typescript-eslint/no-unused-vars
namespace ZChartHelper {
  export const CHART_ID = 911649750;

  export function updateChart(): void {
    const label = `${SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F1')} - ${SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'G1')}`;
    const tenMa = `${SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CHI_TIET_MA, 'F1')}`;
    const HIGH_MA: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'B5');
    const LOW_MA: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'B4');
    const ABS_MA: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'B3');
    const HIGH_VNI: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'C5');
    const LOW_VNI: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'C4');
    const ABS_VNI: number = +SheetHelper.layDuLieuTrongO(SheetHelper.SHEET_CAU_HINH, 'C3');
    const chart = getChartById(CHART_ID, SheetHelper.SHEET_CHI_TIET_MA);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SHEET_CHI_TIET_MA);

    if (!sheet || !chart) {
      console.error(`Sheet hoặc biểu đồ không tồn tại.`);
      return;
    }

    const updatedChart = chart
      .modify()
      .setOption('title', label)
      .setOption('vAxis.minValue', LOW_MA - ABS_MA * 2)
      .setOption('vAxis.maxValue', HIGH_MA + ABS_MA * 2)
      // eslint-disable-next-line @typescript-eslint/naming-convention
      .setOption('series', { 0: { labelInLegend: tenMa }, 1: { labelInLegend: 'VN-INDEX' } })
      // eslint-disable-next-line @typescript-eslint/naming-convention
      .setOption('vAxes', {
        // eslint-disable-next-line @typescript-eslint/naming-convention
        0: { viewWindow: { min: LOW_MA - ABS_MA * 2, max: HIGH_MA + ABS_MA * 2 } },
        // eslint-disable-next-line @typescript-eslint/naming-convention
        1: { viewWindow: { min: LOW_VNI - ABS_VNI * 2, max: HIGH_VNI + ABS_VNI * 2 } }
      })
      .build();

    sheet.updateChart(updatedChart);
  }

  export function createChart(): void {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SHEET_CHI_TIET_MA);
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

  export function getChartById(chartId: number, sheetName: string) {
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

  export function removeChartByID(): void {
    const chart = getChartById(CHART_ID, SheetHelper.SHEET_CHI_TIET_MA);
    if (!chart) {
      console.error('Biểu đồ không tồn tại.');
      return;
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.SHEET_CHI_TIET_MA);
    if (!sheet) {
      console.error(`Sheet ${SheetHelper.SHEET_CHI_TIET_MA} không tồn tại.`);
      return;
    }

    sheet.removeChart(chart);
  }

  export function inRaThongTinChart(): void {
    const spreadsheetId: string = SpreadsheetApp.getActiveSpreadsheet().getId();

    const chartsInfo = Sheets.Spreadsheets?.get(spreadsheetId, {
      ranges: [],
      includeGridData: false,
      fields: 'sheets(charts,properties)'
    });

    for (const sheet of chartsInfo?.sheets || []) {
      if (sheet.properties?.title === SheetHelper.SHEET_CHI_TIET_MA) {
        if (sheet.charts && sheet.charts.length > 0) {
          for (const chart of sheet.charts) {
            Logger.log(JSON.stringify(chart));
          }
        } else {
          Logger.log(`Không tìm thấy biểu đồ nào trên bảng ${SheetHelper.SHEET_CHI_TIET_MA}.`);
        }
        break;
      }
    }
  }
}
