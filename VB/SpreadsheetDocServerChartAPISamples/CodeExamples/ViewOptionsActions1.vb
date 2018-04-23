Imports DevExpress.Spreadsheet
Imports DevExpress.Spreadsheet.Charts
Imports DevExpress.Spreadsheet.Drawings
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetChartAPISamples
    Friend Class ViewOptionsActions
        Private Shared Sub ChangeChartAppearance(ByVal workbook As Workbook)
            '            #Region "#ChangeChartAppearance"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask7")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a chart and specify its location.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.ColumnClustered, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("N17")

            ' Add and format the chart title.
            chart.Title.SetValue("Сountries with the largest forest area")
            chart.Title.Font.Color = Color.FromArgb(&H34, &H5E, &H25)

            ' Set no fill for the plot area.
            chart.PlotArea.Fill.SetNoFill()

            ' Apply the gradient fill to the chart area.
            chart.Fill.SetGradientFill(ShapeGradientType.Linear, Color.FromArgb(&HFD, &HEA, &HDA), Color.FromArgb(&H77, &H93, &H3C))
            Dim gradientFill As ShapeGradientFill = chart.Fill.GradientFill
            gradientFill.Stops.Add(0.78F, Color.FromArgb(&HB7, &HDE, &HE8))
            gradientFill.Angle = 90

            ' Set the picture fill for the data series.
            chart.Series(0).Fill.SetPictureFill("Pictures\PictureFill.png")

            ' Customize the axis appearance.
            Dim axisCollection As AxisCollection = chart.PrimaryAxes
            For Each axis As Axis In axisCollection
                axis.MajorTickMarks = AxisTickMarks.None
                axis.Outline.SetSolidFill(Color.FromArgb(&H34, &H5E, &H25))
                axis.Outline.Width = 1.25
            Next axis
            ' Change the scale of the value axis.
            Dim valueAxis As Axis = axisCollection(1)
            valueAxis.Scaling.AutoMax = False
            valueAxis.Scaling.Max = 8000000
            valueAxis.Scaling.AutoMin = False
            valueAxis.Scaling.Min = 0
            ' Specify display units for the value axis.
            valueAxis.DisplayUnits.UnitType = DisplayUnitType.Thousands
            valueAxis.DisplayUnits.ShowLabel = True

            ' Hide the legend.
            chart.Legend.Visible = False
            '            #End Region ' #ChangeChartAppearance
        End Sub

        Private Shared Sub ApplyGradientToChartBackground(ByVal workbook As Workbook)
            '            #Region "#ApplyGradientToChartBackground"
            Dim worksheet As Worksheet = workbook.Worksheets("chartScatter")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a chart and specify its location.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.ScatterLineMarkers, worksheet("C2:D52"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L17")

            ' Set the series line color.
            chart.Series(0).Outline.SetSolidFill(Color.FromArgb(&HBC, &HCF, &H2))

            ' Specify the data markers.
            Dim markerOptions As Marker = chart.Series(0).Marker
            markerOptions.Symbol = MarkerStyle.Diamond
            markerOptions.Size = 10
            markerOptions.Fill.SetSolidFill(Color.FromArgb(&HBC, &HCF, &H2))
            markerOptions.Outline.SetNoFill()

            ' Set no fill for the plot area.
            chart.PlotArea.Fill.SetNoFill()

            ' Apply the gradient fill to the chart area.
            chart.Fill.SetGradientFill(ShapeGradientType.Circle, Color.FromArgb(&HE3, &H48, &H3), Color.FromArgb(&H0, &H32, &H86))
            Dim gradientFill As ShapeGradientFill = chart.Fill.GradientFill
            gradientFill.FillRect.Left = 0.5
            gradientFill.FillRect.Right = 0.5
            gradientFill.FillRect.Bottom = 0.5
            gradientFill.FillRect.Top = 0.5

            ' Set the X-axis scale.
            Dim axisX As Axis = chart.PrimaryAxes(0)
            axisX.Scaling.AutoMax = False
            axisX.Scaling.AutoMin = False
            axisX.Scaling.Max = 60.0
            axisX.Scaling.Min = -60.0
            axisX.MajorGridlines.Visible = True
            axisX.Visible = False

            ' Set the Y-axis scale.
            Dim axisY As Axis = chart.PrimaryAxes(1)
            axisY.Scaling.AutoMax = False
            axisY.Scaling.AutoMin = False
            axisY.Scaling.Max = 50.0
            axisY.Scaling.Min = -50.0
            axisY.MajorUnit = 10.0
            axisY.Visible = False

            ' Hide the chart legend.
            chart.Legend.Visible = False
            '            #End Region ' #ApplyGradientToChartBackground
        End Sub

        Private Shared Sub CustomWallsAndFloor(ByVal workbook As Workbook)
            '            #Region "#CustomWallsAndFloor"
            Dim worksheet As Worksheet = workbook.Worksheets("chartTask5")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a chart and specify its location.
            Dim chart As Chart = worksheet.Charts.Add(ChartType.Column3DClustered, worksheet("B2:C8"))
            chart.TopLeftCell = worksheet.Cells("F2")
            chart.BottomRightCell = worksheet.Cells("L15")

            ' Specify that each data point in the series has a different color.
            chart.Views(0).VaryColors = True
            ' Specify the series outline.
            chart.Series(0).Outline.SetSolidFill(Color.AntiqueWhite)
            ' Hide the legend.
            chart.Legend.Visible = False

            ' Specify the side wall color.
            chart.View3D.SideWall.Fill.SetSolidFill(Color.FromArgb(&HDC, &HFA, &HDD))
            ' Specify the pattern fill for the back wall.
            chart.View3D.BackWall.Fill.SetPatternFill(Color.FromArgb(&H9C, &HFB, &H9F), Color.WhiteSmoke, DevExpress.Spreadsheet.Drawings.ShapeFillPatternType.DiagonalBrick)

            Dim floorOptions As SurfaceOptions = chart.View3D.Floor
            ' Specify the floor color.
            floorOptions.Fill.SetSolidFill(Color.FromArgb(&HFA, &HDC, &HF9))
            ' Specify the floor border. 
            floorOptions.Outline.SetSolidFill(Color.FromArgb(&HB4, &H95, &HDE))
            floorOptions.Outline.Width = 1.25
            '            #End Region ' #CustomWallsAndFloor
        End Sub
    End Class
End Namespace
