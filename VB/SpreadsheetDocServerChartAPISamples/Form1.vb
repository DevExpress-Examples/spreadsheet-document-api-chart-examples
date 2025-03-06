Imports DevExpress.Spreadsheet
Imports SpreadsheetChartAPIActions
Imports SpreadsheetDocServerChartAPISamples


Namespace SpreadsheetChartAPISamples
    Public Class Form1
        Inherits Form

#Region "#CreateWorkbook"
        ' Create a new Workbook object.
        Private workbook As New Workbook()
#End Region ' #CreateWorkbook

        Private treeList1 As DevExpress.XtraTreeList.TreeList
        Private WithEvents btnOpenExcel As Button
        Private treeListColumn1 As DevExpress.XtraTreeList.Columns.TreeListColumn
        Private splitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl

        Public Sub New()
            InitializeComponent()
            InitTreeListControl()
            workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic
        End Sub

        Private Sub InitTreeListControl()
            Dim examples As New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub

        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
#Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Create Charts"))
            examples.Add(New SpreadsheetNode("Chart Data"))
            examples.Add(New SpreadsheetNode("Data Labels"))
            examples.Add(New SpreadsheetNode("Chart Axes"))
            examples.Add(New SpreadsheetNode("Chart Legends"))
            examples.Add(New SpreadsheetNode("Protection"))
            examples.Add(New SpreadsheetNode("Chart Series"))
            examples.Add(New SpreadsheetNode("Sparklines"))
            examples.Add(New SpreadsheetNode("Chart Styles"))
            examples.Add(New SpreadsheetNode("Chart Titles"))
            examples.Add(New SpreadsheetNode("Trendlines"))
            examples.Add(New SpreadsheetNode("View Options"))
            '			#End Region

            '			#Region "ExampleNodes"
            ' Add nodes to the "Create Charts" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Create Bar Chart", ChartActions.CreateBarChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Bubble Chart", ChartActions.CreateBubbleChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Column Chart", ChartActions.CreateColumnChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Complex Chart", ChartActions.CreateComplexChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Doughnut Chart", ChartActions.CreateDoughnutChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create 3D Pie Chart", ChartActions.CreatePie3dChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Pie Chart", ChartActions.CreatePieChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Pie of Pie Chart", ChartActions.CreatePieOfPieChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Scatter Chart", ChartActions.CreateScatterChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Stock Chart", ChartActions.CreateStockChartAction))
            examples(0).Groups.Add(New SpreadsheetExample("Change Chart Type", ChartActions.ChangeChartTypeAction))

            ' Add nodes to the "Chart Data" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Change Data Reference", CreationAndDataActions.ChangeDataReferenceAction))
            examples(1).Groups.Add(New SpreadsheetExample("Create Chart And Select Data", CreationAndDataActions.CreateChartAndSelectDataAction))
            examples(1).Groups.Add(New SpreadsheetExample("Create Chart And Select Data Direction", CreationAndDataActions.CreateChartAndSelectDataDirectionAction))
            examples(1).Groups.Add(New SpreadsheetExample("Create Chart From Range", CreationAndDataActions.CreateChartFromRangeAction))
            examples(1).Groups.Add(New SpreadsheetExample("Create Chart With Complex Range", CreationAndDataActions.CreateChartWithComplexRangeAction))
            examples(1).Groups.Add(New SpreadsheetExample("Create Chart With Literal Data", CreationAndDataActions.CreateChartWithLiteralDataAction))

            ' Add nodes to the "Data Labels" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Show Data Labels", DataLabelActions.ShowDataLabelsAction))
            examples(2).Groups.Add(New SpreadsheetExample("Set Data Label Position", DataLabelActions.SetDataLabelsPositionAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Labels Per Series", DataLabelActions.DataLabelsPerSeriesAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Labels Per Point", DataLabelActions.DataLabelsPerPointAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Label Number Format", DataLabelActions.DataLabelsNumberFormatAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Label Separator", DataLabelActions.DataLabelsSeparatorAction))

            ' Add nodes to the "Chart Legend" group of examples.
            examples(3).Groups.Add(New SpreadsheetExample("Hide Legend", LegendActions.HideLegendAction))
            examples(3).Groups.Add(New SpreadsheetExample("Set Legend Position", LegendActions.SetLegendPositionAction))
            examples(3).Groups.Add(New SpreadsheetExample("Exclude Legend Entry", LegendActions.ExcludeLegendEntryAction))

            ' Add nodes to the "Protection" group of examples.
            examples(4).Groups.Add(New SpreadsheetExample("Protect the Chart", ProtectionActions.ProtectChartAction))


            ' Add nodes to the "Chart Series" group of examples.
            examples(5).Groups.Add(New SpreadsheetExample("Change Series Type", SeriesActions.ChangeSeriesTypeAction))
            examples(5).Groups.Add(New SpreadsheetExample("Change Series Order", SeriesActions.ChangeSeriesOrderAction))
            examples(5).Groups.Add(New SpreadsheetExample("Change Series Arguments", SeriesActions.ChangeSeriesArgumentsAction))
            examples(5).Groups.Add(New SpreadsheetExample("Use Secondary Axes", SeriesActions.UseSecondaryAxesAction))
            examples(5).Groups.Add(New SpreadsheetExample("Remove Series", SeriesActions.RemoveSeriesAction))

            ' Add nodes to the "Sparklines" group of examples.
            examples(7).Groups.Add(New SpreadsheetExample("Create a Sparkline", SparklineActions.CreateSparklineGroupsAction))
            examples(7).Groups.Add(New SpreadsheetExample("Customize Sparkline Appearance", SparklineActions.CustomizeSparklineAppearanceAction))
            examples(7).Groups.Add(New SpreadsheetExample("Rearrange Sparklines", SparklineActions.RearrangeSparklinesAction))
            examples(7).Groups.Add(New SpreadsheetExample("Specify Sparkline Axis Settings", SparklineActions.SpecifyAxisSettingsAction))

            ' Add nodes to the "Style" group of examples.
            examples(8).Groups.Add(New SpreadsheetExample("Set Chart Style", StyleActions.SetChartStyleAction))
            examples(8).Groups.Add(New SpreadsheetExample("Custom Series Color", StyleActions.CustomSeriesColorAction))
            examples(8).Groups.Add(New SpreadsheetExample("Set Chart Font", StyleActions.SetChartFontAction))
            examples(8).Groups.Add(New SpreadsheetExample("Set Transparency", StyleActions.TransparencyAction))

            ' Add nodes to the "Titles" group of examples.
            examples(9).Groups.Add(New SpreadsheetExample("Set Title Text", TitlesActions.SetChartTitleTextAction))
            examples(9).Groups.Add(New SpreadsheetExample("Link Title to Cell Range", TitlesActions.LinkChartTitleToCellRangeAction))
            examples(9).Groups.Add(New SpreadsheetExample("Show Chart Title", TitlesActions.ShowChartTitleAction))
            examples(9).Groups.Add(New SpreadsheetExample("Set Axis Title", TitlesActions.SetAxisTitleTextAction))
            examples(9).Groups.Add(New SpreadsheetExample("Link Axis Title to Cell Range ", TitlesActions.LinkAxisTitleToCellRangeAction))
            examples(9).Groups.Add(New SpreadsheetExample("Show Axis Title", TitlesActions.ShowAxisTitleAction))


            ' Add nodes to the "Trendlines" group of examples.
            examples(10).Groups.Add(New SpreadsheetExample("Display Trendline", TrendlineActions.TrendlinesAction))
            examples(10).Groups.Add(New SpreadsheetExample("Specify Trendline Label", TrendlineActions.TrendlineLabelAction))
            examples(10).Groups.Add(New SpreadsheetExample("Customize Trendline", TrendlineActions.TrendlineCustomizationAction))

            ' Add nodes to the "View Options" group of examples.
            examples(11).Groups.Add(New SpreadsheetExample("Apply Gradient To a Chart Background", ViewOptionsActions.ApplyGradientToChartBackgroundAction))
            examples(11).Groups.Add(New SpreadsheetExample("Change Chart Appearance", ViewOptionsActions.ChangeChartAppearanceAction))
            examples(11).Groups.Add(New SpreadsheetExample("Custom Walls And Floor", ViewOptionsActions.CustomWallsAndFloorAction))
            examples(11).Groups.Add(New SpreadsheetExample("Specify Gap Width", ViewOptionsActions.GapWidthAction))
            examples(11).Groups.Add(New SpreadsheetExample("Show Automatic Markers", ViewOptionsActions.ShowAutomaticMarkersAction))
            examples(11).Groups.Add(New SpreadsheetExample("Show Custom Markers", ViewOptionsActions.ShowCustomMarkersAction))
            examples(11).Groups.Add(New SpreadsheetExample("Smooth Lines", ViewOptionsActions.SmoothLinesAction))
            examples(11).Groups.Add(New SpreadsheetExample("Vary Colors By Point", ViewOptionsActions.VaryColorsByPointAction))
#End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub


        Private Sub btnOpenExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOpenExcel.Click
            LoadDocumentFromFile()
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then
                Return
            End If
            Dim action As Action(Of Workbook) = example.Action
            action(workbook)
            SaveDocumentToFile()
        End Sub

        ' ------------------- Load and Save a Document -------------------
        Private Sub LoadDocumentFromFile()
#Region "#LoadDocumentFromFile"
            ' Load a workbook from the file.
            workbook.LoadDocument("Document.xlsx", DocumentFormat.OpenXml)
#End Region ' #LoadDocumentFromFile
        End Sub


        Private Sub SaveDocumentToFile()
#Region "#SaveDocumentToFile"
            ' Save the modified document to the file.
            workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml)
#End Region ' #SaveDocumentToFile
            Process.Start(New ProcessStartInfo("SavedDocument.xlsx") With {.UseShellExecute = True})
        End Sub

        Private Sub InitializeComponent()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.treeListColumn1 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
            Me.btnOpenExcel = New System.Windows.Forms.Button()
            Me.splitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            DirectCast(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            DirectCast(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' treeList1
            ' 
            Me.treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            Me.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Blue
            Me.treeList1.Appearance.FocusedCell.Options.UseFont = True
            Me.treeList1.Appearance.FocusedCell.Options.UseForeColor = True
            Me.treeList1.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() {Me.treeListColumn1})
            Me.treeList1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.treeList1.Location = New System.Drawing.Point(0, 0)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.OptionsBehavior.Editable = False
            Me.treeList1.OptionsView.ShowColumns = False
            Me.treeList1.OptionsView.ShowIndicator = False
            Me.treeList1.Size = New System.Drawing.Size(497, 638)
            Me.treeList1.TabIndex = 0
            ' 
            ' treeListColumn1
            ' 
            Me.treeListColumn1.Caption = "Name"
            Me.treeListColumn1.FieldName = "Name"
            Me.treeListColumn1.Name = "treeListColumn1"
            Me.treeListColumn1.Visible = True
            Me.treeListColumn1.VisibleIndex = 0
            Me.treeListColumn1.Width = 92
            ' 
            ' btnOpenExcel
            ' 
            Me.btnOpenExcel.Dock = System.Windows.Forms.DockStyle.Fill
            Me.btnOpenExcel.Location = New System.Drawing.Point(0, 0)
            Me.btnOpenExcel.Name = "button1"
            Me.btnOpenExcel.Size = New System.Drawing.Size(497, 57)
            Me.btnOpenExcel.TabIndex = 1
            Me.btnOpenExcel.Text = "Run"
            Me.btnOpenExcel.UseVisualStyleBackColor = True
            'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
            'ORIGINAL LINE: this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            ' 
            ' splitContainerControl1
            ' 
            Me.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.splitContainerControl1.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.splitContainerControl1.Horizontal = False
            Me.splitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl1.Name = "splitContainerControl1"
            Me.splitContainerControl1.Panel1.Controls.Add(Me.treeList1)
            Me.splitContainerControl1.Panel1.Text = "Panel1"
            Me.splitContainerControl1.Panel2.Controls.Add(Me.btnOpenExcel)
            Me.splitContainerControl1.Panel2.Text = "Panel2"
            Me.splitContainerControl1.Size = New System.Drawing.Size(497, 700)
            Me.splitContainerControl1.SplitterPosition = 57
            Me.splitContainerControl1.TabIndex = 2
            Me.splitContainerControl1.Text = "splitContainerControl1"
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(497, 700)
            Me.Controls.Add(Me.splitContainerControl1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            DirectCast(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            DirectCast(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl1.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub
    End Class
End Namespace
