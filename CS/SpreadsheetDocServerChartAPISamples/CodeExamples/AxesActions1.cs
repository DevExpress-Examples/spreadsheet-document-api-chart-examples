using System;
using System.Drawing;
using System.Globalization;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using DevExpress.Spreadsheet.Drawings;
using DevExpress.Utils;

namespace SpreadsheetChartAPIActions
{
    public static class AxesActions1
    {
        static void DisplayUnits(Workbook workbook)
        {
            #region #DisplayUnits
            Worksheet worksheet = workbook.Worksheets["chartTask7"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a chart and specify its location.
            Chart chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet["B2:C8"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["N17"];

            // Change the scale of the value axis.
            AxisCollection axisCollection = chart.PrimaryAxes;
            Axis valueAxis = axisCollection[1];
            valueAxis.Scaling.AutoMax = false;
            valueAxis.Scaling.Max = 8000000;
            valueAxis.Scaling.AutoMin = false;
            valueAxis.Scaling.Min = 0;

            // Specify display units for the value axis.
            valueAxis.DisplayUnits.UnitType = DisplayUnitType.Thousands;
            valueAxis.DisplayUnits.ShowLabel = true;

            // Set the chart style.
            chart.Style = ChartStyle.ColorBevel;
            chart.Views[0].VaryColors = true;

            // Hide the legend.
            chart.Legend.Visible = false;
            #endregion #DisplayUnits
        }
    }
}
