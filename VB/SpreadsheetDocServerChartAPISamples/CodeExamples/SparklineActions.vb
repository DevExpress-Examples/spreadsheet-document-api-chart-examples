Imports DevExpress.Spreadsheet
Imports System.Collections.Generic
Imports System.Drawing

Namespace SpreadsheetChartAPIActions
    Public NotInheritable Class SparklineActions
        Private Sub New()
        End Sub
        Private Shared Sub CreateSparklineGroups(ByVal workbook As Workbook)
            '			#Region "#CreateSparklineGroups"
            Dim worksheet As Worksheet = workbook.Worksheets("SparklineExamples")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a group of line sparklines.
            Dim quarterlyGroup As SparklineGroup = worksheet.SparklineGroups.Add(worksheet("G4:G6"), worksheet("C4:F4,C5:F5,C6:F6"), SparklineGroupType.Line)
            ' Add one more sparkline to the existing group.
            quarterlyGroup.Sparklines.Add(6, 6, worksheet("C7:F7"))

            ' Display a column sparkline in the total cell.
            Dim totalGroup As SparklineGroup = worksheet.SparklineGroups.Add(worksheet("G8"), worksheet("C8:F8"), SparklineGroupType.Column)
            '			#End Region ' #CreateSparklineGroups
        End Sub

        Private Shared Sub RearrangeSparklines(ByVal workbook As Workbook)
            '			#Region "#RearrangeSparklines"
            Dim worksheet As Worksheet = workbook.Worksheets("SparklineExamples")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a group of line sparklines.
            Dim lineGroup As SparklineGroup = worksheet.SparklineGroups.Add(worksheet("G4:G7"), worksheet("C4:F4,C5:F5,C6:F6, C7:F7"), SparklineGroupType.Line)

            ' Rearrange sparklines by grouping the second and fourth sparklines together and changing the group type to "Column".
            Dim sparklineG5 As Sparkline = lineGroup.Sparklines(1)
            Dim sparklineG7 As Sparkline = lineGroup.Sparklines(3)
            Dim columnGroup As SparklineGroup = worksheet.SparklineGroups.Add(New List(Of Sparkline)(New Sparkline() {sparklineG5, sparklineG7}), SparklineGroupType.Column)
            '			#End Region ' #RearrangeSparklines
        End Sub

        Private Shared Sub CustomizeSparklineAppearance(ByVal workbook As Workbook)
            '			#Region "#CustomizeSparklineAppearance"
            Dim worksheet As Worksheet = workbook.Worksheets("SparklineExamples")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a group of line sparklines.
            Dim lineGroup As SparklineGroup = worksheet.SparklineGroups.Add(worksheet("G4:G7"), worksheet("C4:F4,C5:F5,C6:F6, C7:F7"), SparklineGroupType.Line)

            ' Customize the group appearance.
            ' Set the sparkline color.
            lineGroup.SeriesColor = Color.FromArgb(&H1F, &H49, &H7D)

            ' Set the sparkline weight.
            lineGroup.LineWeight = 1.5

            ' Display data markers on the sparklines and specify their color.
            Dim points As SparklinePoints = lineGroup.Points
            points.Markers.IsVisible = True
            points.Markers.Color = Color.FromArgb(&H4B, &HAC, &HC6)

            ' Highlight the highest and lowest points on each sparkline in the group.
            points.Highest.Color = Color.FromArgb(&HA9, &HD6, &H4F)
            points.Lowest.Color = Color.FromArgb(&H80, &H64, &HA2)
            '			#End Region ' #CustomizeSparklineAppearance
        End Sub

        Private Shared Sub SpecifyAxisSettings(ByVal workbook As Workbook)
            '			#Region "#SpecifyAxisSettings"
            Dim worksheet As Worksheet = workbook.Worksheets("SparklineExamples")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a group of column sparklines.
            Dim columnGroup As SparklineGroup = worksheet.SparklineGroups.Add(worksheet("G4:G7"), worksheet("C4:F4,C5:F5,C6:F6, C7:F7"), SparklineGroupType.Column)

            ' Specify the vertical axis options.
            Dim verticalAxis As SparklineVerticalAxis = columnGroup.VerticalAxis
            ' Set the custom minimum value for the vertical axis.
            verticalAxis.MinScaleType = SparklineAxisScaling.Custom
            verticalAxis.MinCustomValue = 0
            ' Set the custom maximum value for the vertical axis.
            verticalAxis.MaxScaleType = SparklineAxisScaling.Custom
            verticalAxis.MaxCustomValue = 12000
            '			#End Region ' #SpecifyAxisSettings
        End Sub
    End Class
End Namespace
