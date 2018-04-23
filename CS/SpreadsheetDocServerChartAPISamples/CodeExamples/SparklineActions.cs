using DevExpress.Spreadsheet;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetChartAPIActions
{
    public static class SparklineActions
    {
        static void CreateSparklineGroups(Workbook workbook)
        {
            #region #CreateSparklineGroups
            Worksheet worksheet = workbook.Worksheets["SparklineExamples"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a group of line sparklines.
            SparklineGroup quarterlyGroup = worksheet.SparklineGroups.Add(worksheet["G4:G6"], worksheet["C4:F4,C5:F5,C6:F6"], SparklineGroupType.Line);
            // Add one more sparkline to the existing group.
            quarterlyGroup.Sparklines.Add(6, 6, worksheet["C7:F7"]);

            // Display a column sparkline in the total cell.
            SparklineGroup totalGroup = worksheet.SparklineGroups.Add(worksheet["G8"], worksheet["C8:F8"], SparklineGroupType.Column);
            #endregion #CreateSparklineGroups
        }

        static void RearrangeSparklines(Workbook workbook)
        {
            #region #RearrangeSparklines
            Worksheet worksheet = workbook.Worksheets["SparklineExamples"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a group of line sparklines.
            SparklineGroup lineGroup = worksheet.SparklineGroups.Add(worksheet["G4:G7"], worksheet["C4:F4,C5:F5,C6:F6, C7:F7"], SparklineGroupType.Line);

            // Rearrange sparklines by grouping the second and fourth sparklines together and changing the group type to "Column".
            Sparkline sparklineG5 = lineGroup.Sparklines[1];
            Sparkline sparklineG7 = lineGroup.Sparklines[3];
            SparklineGroup columnGroup = worksheet.SparklineGroups.Add(new List<Sparkline> { sparklineG5, sparklineG7 }, SparklineGroupType.Column);
            #endregion #RearrangeSparklines
        }

        static void CustomizeSparklineAppearance(Workbook workbook)
        {
            #region #CustomizeSparklineAppearance
            Worksheet worksheet = workbook.Worksheets["SparklineExamples"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a group of line sparklines.
            SparklineGroup lineGroup = worksheet.SparklineGroups.Add(worksheet["G4:G7"], worksheet["C4:F4,C5:F5,C6:F6, C7:F7"], SparklineGroupType.Line);

            // Customize the group appearance.
            // Set the sparkline color.
            lineGroup.SeriesColor = Color.FromArgb(0x1F, 0x49, 0x7D);

            // Set the sparkline weight.
            lineGroup.LineWeight = 1.5;

            // Display data markers on the sparklines and specify their color.
            SparklinePoints points = lineGroup.Points;
            points.Markers.IsVisible = true;
            points.Markers.Color = Color.FromArgb(0x4B, 0xAC, 0xC6);

            // Highlight the highest and lowest points on each sparkline in the group.
            points.Highest.Color = Color.FromArgb(0xA9, 0xD6, 0x4F);
            points.Lowest.Color = Color.FromArgb(0x80, 0x64, 0xA2);
            #endregion #CustomizeSparklineAppearance
        }

        static void SpecifyAxisSettings(Workbook workbook)
        {
            #region #SpecifyAxisSettings
            Worksheet worksheet = workbook.Worksheets["SparklineExamples"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a group of column sparklines.
            SparklineGroup columnGroup = worksheet.SparklineGroups.Add(worksheet["G4:G7"], worksheet["C4:F4,C5:F5,C6:F6, C7:F7"], SparklineGroupType.Column);

            // Specify the vertical axis options.
            SparklineVerticalAxis verticalAxis = columnGroup.VerticalAxis;
            // Set the custom minimum value for the vertical axis.
            verticalAxis.MinScaleType = SparklineAxisScaling.Custom;
            verticalAxis.MinCustomValue = 0;
            // Set the custom maximum value for the vertical axis.
            verticalAxis.MaxScaleType = SparklineAxisScaling.Custom;
            verticalAxis.MaxCustomValue = 12000;
            #endregion #SpecifyAxisSettings
        }
    }
}
