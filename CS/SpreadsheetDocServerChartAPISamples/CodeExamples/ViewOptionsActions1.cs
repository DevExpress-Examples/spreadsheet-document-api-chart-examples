using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using DevExpress.Spreadsheet.Drawings;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetChartAPISamples
{
    class ViewOptionsActions
    {
        static void ChangeChartAppearance(Workbook workbook)
        {
            #region #ChangeChartAppearance
            Worksheet worksheet = workbook.Worksheets["chartTask7"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a chart and specify its location.
            Chart chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet["B2:C8"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["N17"];

            // Add and format the chart title.
            chart.Title.SetValue("Сountries with the largest forest area");
            chart.Title.Font.Color = Color.FromArgb(0x34, 0x5E, 0x25);

            // Set no fill for the plot area.
            chart.PlotArea.Fill.SetNoFill();

            // Apply the gradient fill to the chart area.
            chart.Fill.SetGradientFill(ShapeGradientType.Linear, Color.FromArgb(0xFD, 0xEA, 0xDA), Color.FromArgb(0x77, 0x93, 0x3C));
            ShapeGradientFill gradientFill = chart.Fill.GradientFill;
            gradientFill.Stops.Add(0.78f, Color.FromArgb(0xB7, 0xDE, 0xE8));
            gradientFill.Angle = 90;

            // Set the picture fill for the data series.
            chart.Series[0].Fill.SetPictureFill("Pictures\\PictureFill.png");

            // Customize the axis appearance.
            AxisCollection axisCollection = chart.PrimaryAxes;
            foreach (Axis axis in axisCollection)
            {
                axis.MajorTickMarks = AxisTickMarks.None;
                axis.Outline.SetSolidFill(Color.FromArgb(0x34, 0x5E, 0x25));
                axis.Outline.Width = 1.25;
            }
            // Change the scale of the value axis.
            Axis valueAxis = axisCollection[1];
            valueAxis.Scaling.AutoMax = false;
            valueAxis.Scaling.Max = 8000000;
            valueAxis.Scaling.AutoMin = false;
            valueAxis.Scaling.Min = 0;
            // Specify display units for the value axis.
            valueAxis.DisplayUnits.UnitType = DisplayUnitType.Thousands;
            valueAxis.DisplayUnits.ShowLabel = true;

            // Hide the legend.
            chart.Legend.Visible = false;
            #endregion #ChangeChartAppearance
        }

        static void ApplyGradientToChartBackground(Workbook workbook)
        {
            #region #ApplyGradientToChartBackground
            Worksheet worksheet = workbook.Worksheets["chartScatter"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a chart and specify its location.
            Chart chart = worksheet.Charts.Add(ChartType.ScatterLineMarkers, worksheet["C2:D52"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["L17"];

            // Set the series line color.
            chart.Series[0].Outline.SetSolidFill(Color.FromArgb(0xBC, 0xCF, 0x02));

            // Specify the data markers.
            Marker markerOptions = chart.Series[0].Marker;
            markerOptions.Symbol = MarkerStyle.Diamond;
            markerOptions.Size = 10;
            markerOptions.Fill.SetSolidFill(Color.FromArgb(0xBC, 0xCF, 0x02));
            markerOptions.Outline.SetNoFill();

            // Set no fill for the plot area.
            chart.PlotArea.Fill.SetNoFill();

            // Apply the gradient fill to the chart area.
            chart.Fill.SetGradientFill(ShapeGradientType.Circle, Color.FromArgb(0xE3, 0x48, 0x03), Color.FromArgb(0x00, 0x32, 0x86));
            ShapeGradientFill gradientFill = chart.Fill.GradientFill;
            gradientFill.FillRect.Left = 0.5;
            gradientFill.FillRect.Right = 0.5;
            gradientFill.FillRect.Bottom = 0.5;
            gradientFill.FillRect.Top = 0.5;

            // Set the X-axis scale.
            Axis axisX = chart.PrimaryAxes[0];
            axisX.Scaling.AutoMax = false;
            axisX.Scaling.AutoMin = false;
            axisX.Scaling.Max = 60.0;
            axisX.Scaling.Min = -60.0;
            axisX.MajorGridlines.Visible = true;
            axisX.Visible = false;

            // Set the Y-axis scale.
            Axis axisY = chart.PrimaryAxes[1];
            axisY.Scaling.AutoMax = false;
            axisY.Scaling.AutoMin = false;
            axisY.Scaling.Max = 50.0;
            axisY.Scaling.Min = -50.0;
            axisY.MajorUnit = 10.0;
            axisY.Visible = false;

            // Hide the chart legend.
            chart.Legend.Visible = false;
            #endregion #ApplyGradientToChartBackground
        }

        static void CustomWallsAndFloor(Workbook workbook)
        {
            #region #CustomWallsAndFloor
            Worksheet worksheet = workbook.Worksheets["chartTask5"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a chart and specify its location.
            Chart chart = worksheet.Charts.Add(ChartType.Column3DClustered, worksheet["B2:C8"]);
            chart.TopLeftCell = worksheet.Cells["F2"];
            chart.BottomRightCell = worksheet.Cells["L15"];

            // Specify that each data point in the series has a different color.
            chart.Views[0].VaryColors = true;
            // Specify the series outline.
            chart.Series[0].Outline.SetSolidFill(Color.AntiqueWhite);
            // Hide the legend.
            chart.Legend.Visible = false;

            // Specify the side wall color.
            chart.View3D.SideWall.Fill.SetSolidFill(Color.FromArgb(0xDC, 0xFA, 0xDD));
            // Specify the pattern fill for the back wall.
            chart.View3D.BackWall.Fill.SetPatternFill(Color.FromArgb(0x9C, 0xFB, 0x9F), Color.WhiteSmoke, DevExpress.Spreadsheet.Drawings.ShapeFillPatternType.DiagonalBrick);

            SurfaceOptions floorOptions = chart.View3D.Floor;
            // Specify the floor color.
            floorOptions.Fill.SetSolidFill(Color.FromArgb(0xFA, 0xDC, 0xF9));
            // Specify the floor border. 
            floorOptions.Outline.SetSolidFill(Color.FromArgb(0xB4, 0x95, 0xDE));
            floorOptions.Outline.Width = 1.25;
            #endregion #CustomWallsAndFloor
        }
    }
}
