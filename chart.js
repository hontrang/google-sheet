const CHART_ID = 559458186;
const API_KEY = "AIzaSyBr-2mGf58LY2kXma1KFUFpgEGuI8lNhgw";

function updateChart() {
  const tenMa = `${SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1")} - ${SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "G1")}`;
  const HIGH = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD51");
  const LOW = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD50");
  const ABS = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD49");
  const chart = getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);

  if (!chart) {
    console.error('Không tìm thấy biểu đồ.');
    return;
  }

  const updatedChart = chart.modify()
    .setOption('title', tenMa)
    .setOption('vAxis.minValue', LOW - ABS * 2)
    .setOption('vAxis.maxValue', HIGH + ABS * 2)
    .setOption('series', { 0: { labelInLegend: tenMa } })
    .setOption('vAxes', {
      0: { viewWindow: { min: LOW - ABS * 2, max: HIGH + ABS * 2 }},
      1: { viewWindow: { min: 1100, max: 1250 }}
    })
    .build();

  sheet.updateChart(updatedChart);
}

function createChart() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);
  if (!sheet) {
    console.error('Không tìm thấy sheet.');
    return;
  }

  const range = sheet.getRange("A1:B8");
  const chartBuilder = sheet.newChart()
    .addRange(range)
    .setChartType(Charts.ChartType.LINE)
    .setPosition(10, 10, 0, 0)
    .setOption('title', 'My Line Chart!');

  const chart = chartBuilder.build();
  sheet.insertChart(chart);
  console.log(chart.getChartId());
}

function getChartById(chartId, sheetName) {
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

function removeChartByID() {
  const chart = getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
  if (!chart) {
    console.error('Biểu đồ không tồn tại.');
    return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);
  if (!sheet) {
    console.error(`Sheet ${SheetUtility.SHEET_CHI_TIET_MA} không tồn tại.`);
    return;
  }

  sheet.removeChart(chart);
}

function inRaThongTinChart() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Gọi API để truy xuất thông tin về biểu đồ
  const chartsInfo = Sheets.Spreadsheets.get(spreadsheetId, {
    ranges: [],
    includeGridData: false,
    fields: "sheets(charts,properties)"
  });

  // Duyệt qua tất cả các bảng
  for (const sheet of chartsInfo.sheets) {
    if (sheet.properties.title === SheetUtility.SHEET_CHI_TIET_MA) {
      // Kiểm tra xem bảng có biểu đồ không
      if (sheet.charts && sheet.charts.length > 0) {
        for (const chart of sheet.charts) {
          Logger.log(JSON.stringify(chart));
        }
      } else {
        Logger.log(`Không tìm thấy biểu đồ nào trên bảng ${SheetUtility.SHEET_CHI_TIET_MA}.`);
      }
      break; // Thoát khỏi vòng lặp sau khi đã xử lý bảng tương ứng
    }
  }
}
