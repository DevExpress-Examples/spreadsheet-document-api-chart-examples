Imports System
Imports System.Drawing
Imports System.Globalization
Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports DevExpress.Spreadsheet.Drawings
Imports DevExpress.Utils

Namespace SpreadsheetChartAPIActions
    Public NotInheritable Class AxesActions1

        Private Sub New()
        End Sub

        Private Shared Sub DisplayUnits(ByVal workbook As Workbook)
            '            #Region "#DisplayUnits"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask7")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a chart and specify its location.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("N17")

            ' Change the scale of the value axis.
            Dim axisCollection As AxisCollection = chart.PrimaryAxes
            Dim valueAxis As Axis = axisCollection(1)
            valueAxis.Scaling.AutoMax = False
            valueAxis.Scaling.Max = 8000000
            valueAxis.Scaling.AutoMin = False
            valueAxis.Scaling.Min = 0

            ' Specify display units for the value axis.
            valueAxis.DisplayUnits.UnitType = DisplayUnitType.Thousands
            valueAxis.DisplayUnits.ShowLabel = True

            ' Set the chart style.
            chart.Style = ChartStyle.ColorBevel
            chart.Views(0).VaryColors = True

            ' Hide the legend.
            chart.Legend.Visible = False
            '            #End Region ' #DisplayUnits
        End Sub
    End Class
End Namespace
