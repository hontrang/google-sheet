/* eslint-disable @typescript-eslint/no-extraneous-class */
import { SheetHelper } from '@src/utility/SheetHelper';

export class ZchartHelper {
  public static updateChart(): void {
    const sheetHelper = new SheetHelper();
    const tenMa = sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'B1');
    const title = `*${tenMa}* - ${sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetChiTietMa, 'D1')}`;
    const HIGH_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B5');
    const LOW_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B4');
    const ABS_MA: number = +sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'B3');
    const HIGH_VNI: number = +sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'D5');
    const LOW_VNI: number = +sheetHelper.layDuLieuTrongO(SheetHelper.sheetName.sheetCauHinh, 'D4');
    let chart = this.getChart(SheetHelper.sheetName.sheetChiTietMa);
    const sheet = SpreadsheetApp.getActive().getSheetByName(SheetHelper.sheetName.sheetChiTietMa);

    if (!sheet || chart === null) {
      console.error(`Sheet hoặc biểu đồ không tồn tại.`);
      return;
    }
    sheet.removeChart(chart);
    chart = sheet.newChart()
      .asLineChart()
      .setRange(LOW_MA - ABS_MA * 2, HIGH_MA + ABS_MA * 2)
      .addRange(sheet.getRange('A16:B35'))
      .addRange(sheet.getRange('\'dữ liệu\'!AD18:AE37'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(0)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('curveType', 'function')
      .setOption('title', title)
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#1a1a1a')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.gridlines.count', 10)
      .setOption('hAxis.slantedText', true)
      .setOption('hAxis.slantedTextAngle', 30)
      .setOption('hAxis.textStyle.fontSize', 10)
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('series.0.labelInLegend', tenMa)
      .setOption('trendlines.0.type', 'linear')
      .setOption('vAxes.1.formatOptions.scaleFactor', 100)
      .setOption('vAxes.1.viewWindow.max', HIGH_VNI)
      .setOption('vAxes.1.viewWindow.min', LOW_VNI)
      .setOption('vAxes.1.textStyle.color', '#000000')
      .setOption('series.1.targetAxisIndex', 1)
      .setOption('series.1.labelInLegend', 'VNINDEX')
      .setOption('trendlines.1.type', 'linear')
      .setOption('height', 245)
      .setOption('width', 430)
      .setPosition(13, 1, 0, 0)
      .build();
    sheet.insertChart(chart);
  }

  public static getChart(sheetName: string) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      console.error(`Sheet ${sheetName} không tồn tại.`);
      return null;
    }

    const charts = sheet.getCharts();
    for (const chart of charts) {
      if (chart.getOptions().get('title').toString().startsWith('*')) {
        return chart;
      }
    }

    console.log(`Không tìm thấy biểu đồ trong sheet ${sheetName}.`);
    return null;
  }
}
