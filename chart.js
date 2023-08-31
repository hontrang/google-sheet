const CHART_ID = 559458186;
const API_KEY = "AIzaSyBr-2mGf58LY2kXma1KFUFpgEGuI8lNhgw";
function updateChart() {
  const tenMa = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "F1") + " - " + SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_CHI_TIET_MA, "G1");
  const HIGH = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD51");
  const LOW = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD50");
  const ABS = SheetUtility.layDuLieuTrongO(SheetUtility.SHEET_DU_LIEU, "AD49");
  const chart = getChartById(CHART_ID, SheetUtility.SHEET_CHI_TIET_MA);
  const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtility.SHEET_CHI_TIET_MA);

  var updatedChart = chart.modify()
    .setOption('title', `${tenMa}`)
    .setOption('vAxis.minValue', LOW - ABS * 2)
    .setOption('vAxis.maxValue', HIGH + ABS * 2)
    .setOption('series', {
      0: {
        labelInLegend: tenMa
      }
    })
    .setOption('vAxes', {
      0: {
        viewWindow: {
          min: LOW - ABS * 2, // Đặt giá trị tối thiểu cho trục dọc bên trái
          max: HIGH + ABS * 2 // Đặt giá trị tối đa cho trục dọc bên trái
        }
      }, // Trục dọc bên trái
      1: { // Trục dọc bên phải
        viewWindow: {
          min: 1150, // Đặt giá trị tối thiểu cho trục dọc bên phải
          max: 1300 // Đặt giá trị tối đa cho trục dọc bên phải
        }
      }
    })
    .build();
  // Thay thế biểu đồ cũ bằng biểu đồ đã cập nhật
  sheet.updateChart(updatedChart);
  Logger.log(chart.getChartId());
}

function updateChartAxis() {
    var spreadsheetId = "1GyxkNiyXantim6R6otooAGd6SLoR4J9Db8nCCx216og";
    var sheetName = SheetUtility.SHEET_CHI_TIET_MA; // Thay thế bằng tên sheet chứa biểu đồ của bạn

    // Lấy danh sách tất cả các biểu đồ trong bảng tính
    var charts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getCharts();

    // Tìm biểu đồ dựa trên tiêu đề
    var chartId= CHART_ID;
  

    if (chartId) {
        // Cập nhật giá trị trục dọc bên phải của biểu đồ
        var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + ':batchUpdate';
        var headers = {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
            'Content-Type': 'application/json'
        };

        var payload = {
            "requests": [{
                "updateChartSpec": {
                    "chartId": chartId,
                    "spec": {
                        "basicChart": {
                            "chartType": "LINE",
                            "axis": [{
                                "position": "RIGHT_AXIS",
                                "viewWindowOptions": {
                                    "viewWindowMin": 1150,
                                    "viewWindowMax": 1300
                                }
                            }]
                        }
                    }
                }
            }]
        };

        var options = {
            "method": "POST",
            "headers": headers,
            "payload": JSON.stringify(payload)
        };

        UrlFetchApp.fetch(url, options);
    } else {
        Logger.log("Biểu đồ với tiêu đề '" + chartTitle + "' không được tìm thấy.");
    }
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

function inRaThongTinChart() {
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // Gọi API để truy xuất thông tin về biểu đồ
  var chartsInfo = Sheets.Spreadsheets.get(spreadsheetId, {
    ranges: [],
    includeGridData: false,
    fields: "sheets(charts,properties)"
  });

  // Duyệt qua tất cả các bảng và tìm bảng có tên đúng như cung cấp
  var sheets = chartsInfo.sheets;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].properties.title === SheetUtility.SHEET_CHI_TIET_MA) {
      var charts = sheets[i].charts;
      if (charts) { // Kiểm tra xem bảng có biểu đồ không
        for (var j = 0; j < charts.length; j++) {
          Logger.log(JSON.stringify(charts[j]));
        }
      } else {
        Logger.log("Không tìm thấy biểu đồ nào trên bảng " + SheetUtility.SHEET_CHI_TIET_MA);
      }
      break;
    }
  }
}
