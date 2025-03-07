Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetChartAPISamples
    Public NotInheritable Class ProtectionActions

        Private Sub New()
        End Sub

        Public Shared ProtectChartAction As Action(Of Workbook) = AddressOf ProtectChart

        Private Shared Sub ProtectChart(ByVal workbook As Workbook)
#Region "#ProtectChart"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask3")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a chart and specify its location.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet("B2:D4"))
            chart.TopLeftCell = worksheet.Cells("H2")
            chart.BottomRightCell = worksheet.Cells("N14")

            ' Specify the chart style.
            chart.Style = ChartStyle.ColorDark

            ' Apply the chart protection.
            chart.Options.Protection = ChartProtection.All
#End Region ' #ProtectChart
        End Sub

    End Class
End Namespace
