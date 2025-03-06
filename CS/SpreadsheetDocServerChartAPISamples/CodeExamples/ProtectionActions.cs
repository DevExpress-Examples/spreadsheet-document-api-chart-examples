using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using System;

namespace SpreadsheetChartAPIActions
{
    public static class ProtectionActions
    {
        public static Action<Workbook> ProtectChartAction = ProtectChart;

        static void ProtectChart(Workbook workbook)
        {
            #region #ProtectChart
            Worksheet worksheet = workbook.Worksheets["chartTask3"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a chart and specify its location.
            Chart chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet["B2:D4"]);
            chart.TopLeftCell = worksheet.Cells["H2"];
            chart.BottomRightCell = worksheet.Cells["N14"];

            // Specify the chart style.
            chart.Style = ChartStyle.ColorDark;

            // Apply the chart protection.
            chart.Options.Protection = ChartProtection.All;

            #endregion #ProtectChart
        }

    }
}
