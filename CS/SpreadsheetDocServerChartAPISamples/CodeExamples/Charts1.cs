using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetChartAPISamples
{
    public static class Charts
    {
        static void PieOfPieChart(Workbook workbook)
        {
            #region #PieOfPieChart
            Worksheet worksheet = workbook.Worksheets["chartTask6"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a Pie of Pie chart and specify its position.
            Chart chart = worksheet.Charts.Add(ChartType.PieOfPie, worksheet["B2:C11"]);
            chart.TopLeftCell = worksheet.Cells["E2"];
            chart.BottomRightCell = worksheet.Cells["L16"];

            // Specify the number of data points to be displayed in the secondary chart (the last four values).
            chart.Views[0].SplitType = OfPieSplitType.Position;
            chart.Views[0].SplitPosition = 4;

            // Show data labels as percentage values.
            chart.Views[0].DataLabels.ShowPercent = true;
            #endregion #PieOfPieChart
        }
    }
}
