const CHART_ID = 2048274768;
function updateChart() {
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1");
  const HIGH = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD51");
  const LOW = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD50");
  const ABS = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD49");
  const chart = getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);
  var updatedChart = chart.modify()
    .setOption('title', `${tenMa}`)
    .setOption('vAxis.minValue', LOW - ABS*2)
    .setOption('vAxis.maxValue', HIGH + ABS*2)
    .setOption('series', {
      0: {
        labelInLegend: tenMa
      }
    })
    .build();
  // Thay thế biểu đồ cũ bằng biểu đồ đã cập nhật
  sheet.updateChart(updatedChart);
  Logger.log(chart.getChartId());
}

function createChart() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);
  var chartBuilder = sheet.newChart();
  var range = sheet.getRange("A1:B8");
  chartBuilder.addRange(range)
    .setChartType(Charts.ChartType.LINE)
    .setPosition(10, 10, 0, 0)
    .setOption('title', 'My Line Chart!');
  var char = chartBuilder.build();
  sheet.insertChart(char);
  console.log(char.getChartId());
}
function getChartById(chartId, sheetName) {
  // Lấy tất cả các trang tính
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  // Lấy tất cả các biểu đồ trong trang tính
  var charts = sheet.getCharts();

  // Lặp qua tất cả các biểu đồ
  for (var i = 0; i < charts.length; i++) {
    var chart = charts[i];

    // So sánh ID của biểu đồ với ID mong muốn
    console.log(chart.getChartId() + " " + chart.getOptions().get('title'));
    if (chart.getChartId() === chartId) {
      // Nếu ID khớp, trả về biểu đồ
      return chart;
    }
  }
  // Nếu không tìm thấy biểu đồ nào có ID khớp, trả về null
  return null;
}

function removeChartByID() {
  const chart = getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);
  sheet.removeChart(chart);
}

function listChart() {
  getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
}
