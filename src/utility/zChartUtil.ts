class ZChartUtil {
    static CHART_ID = 911649750;

    static updateChart(): void {
        const label: string = `${SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "F1")} - ${SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "G1")}`;
        const tenMa: string = `${SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CHI_TIET_MA, "F1")}`;
        const HIGH_MA: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "B5");
        const LOW_MA: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "B4");
        const ABS_MA: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "B3");
        const HIGH_VNI: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "C5");
        const LOW_VNI: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "C4");
        const ABS_VNI: number = +SheetUtil.layDuLieuTrongO(SheetUtil.SHEET_CAU_HINH, "C3");

        const chart = ZChartUtil.getChartById(ZChartUtil.CHART_ID, SheetUtil.SHEET_CHI_TIET_MA);
        const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtil.SHEET_CHI_TIET_MA);

        if (!sheet || !chart) {
            console.error(`Sheet hoặc biểu đồ không tồn tại.`);
            return;
        }

        const updatedChart = chart.modify()
            .setOption('title', label)
            .setOption('vAxis.minValue', LOW_MA - ABS_MA * 2)
            .setOption('vAxis.maxValue', HIGH_MA + ABS_MA * 2)
            .setOption('series', {
                0: { labelInLegend: tenMa },
                1: { labelInLegend: "VN-INDEX" }
            })
            .setOption('vAxes', {
                0: { viewWindow: { min: LOW_MA - ABS_MA * 2, max: HIGH_MA + ABS_MA * 2 } },
                1: { viewWindow: { min: LOW_VNI - ABS_VNI * 2, max: HIGH_VNI + ABS_VNI * 2 } }
            }).build();

        sheet.updateChart(updatedChart);
    }

    static createChart(): void {
        const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtil.SHEET_CHI_TIET_MA);
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

    static getChartById(chartId: number, sheetName: string) {
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

    static removeChartByID(): void {
        const chart = ZChartUtil.getChartById(ZChartUtil.CHART_ID, SheetUtil.SHEET_CHI_TIET_MA);
        if (!chart) {
            console.error('Biểu đồ không tồn tại.');
            return;
        }

        const sheet = SpreadsheetApp.getActive().getSheetByName(SheetUtil.SHEET_CHI_TIET_MA);
        if (!sheet) {
            console.error(`Sheet ${SheetUtil.SHEET_CHI_TIET_MA} không tồn tại.`);
            return;
        }

        sheet.removeChart(chart);
    }

    static inRaThongTinChart(): void {
        const spreadsheetId: string = SpreadsheetApp.getActiveSpreadsheet().getId();

        const chartsInfo = Sheets.Spreadsheets?.get(spreadsheetId, {
            ranges: [],
            includeGridData: false,
            fields: "sheets(charts,properties)"
        });

        for (const sheet of (chartsInfo?.sheets || [])) {
            if (sheet.properties?.title === SheetUtil.SHEET_CHI_TIET_MA) {
                if (sheet.charts && sheet.charts.length > 0) {
                    for (const chart of sheet.charts) {
                        Logger.log(JSON.stringify(chart));
                    }
                } else {
                    Logger.log(`Không tìm thấy biểu đồ nào trên bảng ${SheetUtil.SHEET_CHI_TIET_MA}.`);
                }
                break;
            }
        }
    }
}