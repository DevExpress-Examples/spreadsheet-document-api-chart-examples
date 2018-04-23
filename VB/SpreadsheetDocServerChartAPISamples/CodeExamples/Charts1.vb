Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetChartAPISamples
    Public NotInheritable Class Charts

        Private Sub New()
        End Sub

        Private Shared Sub PieOfPieChart(ByVal workbook As Workbook)
            '            #Region "#PieOfPieChart"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask6")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a Pie of Pie chart and specify its position.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.PieOfPie, worksheet("B2:C11"))
            chart.TopLeftCell = worksheet.Cells("E2")
            chart.BottomRightCell = worksheet.Cells("L16")

            ' Specify the number of data points to be displayed in the secondary chart (the last four values).
            chart.Views(0).SplitType = OfPieSplitType.Position
            chart.Views(0).SplitPosition = 4

            ' Show data labels as percentage values.
            chart.Views(0).DataLabels.ShowPercent = True
            '            #End Region ' #PieOfPieChart
        End Sub
    End Class
End Namespace
