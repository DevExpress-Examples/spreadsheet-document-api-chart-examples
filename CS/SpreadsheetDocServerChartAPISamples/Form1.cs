using DevExpress.Spreadsheet;
using SpreadsheetChartAPIActions;
using SpreadsheetDocServerChartAPISamples;
using System;
using System.Diagnostics;
using System.Windows.Forms;


namespace SpreadsheetChartAPISamples
{
    public class Form1 : Form
    {
        #region #CreateWorkbook
        // Create a new Workbook object.
        Workbook workbook = new Workbook();
        #endregion #CreateWorkbook

        private DevExpress.XtraTreeList.TreeList treeList1;
        private System.Windows.Forms.Button btnOpenExcel;
        private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn1;
        private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;

        public Form1()
        {
            InitializeComponent();
            InitTreeListControl();
            workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic;
        }

        void InitTreeListControl()
        {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }

        void InitData(GroupsOfSpreadsheetExamples examples)
        {
            #region GroupNodes        
            examples.Add(new SpreadsheetNode("Chart Axes"));
            examples.Add(new SpreadsheetNode("Create Charts"));
            examples.Add(new SpreadsheetNode("Chart Data"));
            examples.Add(new SpreadsheetNode("Data Labels"));
            examples.Add(new SpreadsheetNode("Chart Legends"));
            examples.Add(new SpreadsheetNode("Protection"));
            examples.Add(new SpreadsheetNode("Chart Series"));
            examples.Add(new SpreadsheetNode("Sparklines"));
            examples.Add(new SpreadsheetNode("Chart Styles"));
            examples.Add(new SpreadsheetNode("Chart Titles"));
            examples.Add(new SpreadsheetNode("Trendlines"));
            examples.Add(new SpreadsheetNode("View Options"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Axes" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Min and Max Values", AxesActions.MinAndMaxValuesAction));
            examples[0].Groups.Add(new SpreadsheetExample("Major Units", AxesActions.MajorUnitsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Major and Minor Gridlines", AxesActions.MajorAndMinorGridlinesAction));
            examples[0].Groups.Add(new SpreadsheetExample("Labels Number Format", AxesActions.LabelsNumberFormatAction));
            examples[0].Groups.Add(new SpreadsheetExample("Hide Tick Marks", AxesActions.HideTickMarksAction));
            examples[0].Groups.Add(new SpreadsheetExample("Hide Axis Line", AxesActions.HideAxisLineAction));
            examples[0].Groups.Add(new SpreadsheetExample("Axis Position", AxesActions.PositionAction));
            examples[0].Groups.Add(new SpreadsheetExample("Axis Orientation", AxesActions.OrientationAction));
            examples[0].Groups.Add(new SpreadsheetExample("Log Scale", AxesActions.LogScaleAction));
            examples[0].Groups.Add(new SpreadsheetExample("Display Units", AxesActions.DisplayUnitsAction));


            // Add nodes to the "Create Charts" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Create Bar Chart", ChartsActions.CreateBarChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Bubble Chart", ChartsActions.CreateBubbleChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Column Chart", ChartsActions.CreateColumnChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Complex Chart", ChartsActions.CreateComplexChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Doughnut Chart", ChartsActions.CreateDoughnutChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create 3D Pie Chart", ChartsActions.CreatePie3dChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Pie Chart", ChartsActions.CreatePieChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Pie of Pie Chart", ChartsActions.CreatePieOfPieChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Scatter Chart", ChartsActions.CreateScatterChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Create Stock Chart", ChartsActions.CreateStockChartAction));
            examples[1].Groups.Add(new SpreadsheetExample("Change Chart Type", ChartsActions.ChangeChartTypeAction));

            // Add nodes to the "Chart Data" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Change Data Reference", CreationAndDataActions.ChangeDataReferenceAction));
            examples[2].Groups.Add(new SpreadsheetExample("Create Chart And Select Data", CreationAndDataActions.CreateChartAndSelectDataAction));
            examples[2].Groups.Add(new SpreadsheetExample("Create Chart And Select Data Direction", CreationAndDataActions.CreateChartAndSelectDataDirectionAction));
            examples[2].Groups.Add(new SpreadsheetExample("Create Chart From Range", CreationAndDataActions.CreateChartFromRangeAction));
            examples[2].Groups.Add(new SpreadsheetExample("Create Chart With Complex Range", CreationAndDataActions.CreateChartWithComplexRangeAction));
            examples[2].Groups.Add(new SpreadsheetExample("Create Chart With Literal Data", CreationAndDataActions.CreateChartWithLiteralDataAction));

            // Add nodes to the "Data Labels" group of examples.
            examples[3].Groups.Add(new SpreadsheetExample("Show Data Labels", DataLabelsActions.ShowDataLabelsAction));
            examples[3].Groups.Add(new SpreadsheetExample("Set Data Label Position", DataLabelsActions.SetDataLabelsPositionAction));
            examples[3].Groups.Add(new SpreadsheetExample("Data Labels Per Series", DataLabelsActions.DataLabelsPerSeriesAction));
            examples[3].Groups.Add(new SpreadsheetExample("Data Labels Per Point", DataLabelsActions.DataLabelsPerPointAction));
            examples[3].Groups.Add(new SpreadsheetExample("Data Label Number Format", DataLabelsActions.DataLabelsNumberFormatAction));
            examples[3].Groups.Add(new SpreadsheetExample("Data Label Separator", DataLabelsActions.DataLabelsSeparatorAction));

            // Add nodes to the "Chart Legend" group of examples.
            examples[4].Groups.Add(new SpreadsheetExample("Hide Legend", LegendActions.HideLegendAction));
            examples[4].Groups.Add(new SpreadsheetExample("Set Legend Position", LegendActions.SetLegendPositionAction));
            examples[4].Groups.Add(new SpreadsheetExample("Exclude Legend Entry", LegendActions.ExcludeLegendEntryAction));

            // Add nodes to the "Protection" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Protect the Chart", ProtectionActions.ProtectChartAction));

            // Add nodes to the "Chart Series" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Change Series Type", SeriesActions.ChangeSeriesTypeAction));
            examples[6].Groups.Add(new SpreadsheetExample("Change Series Order", SeriesActions.ChangeSeriesOrderAction));
            examples[6].Groups.Add(new SpreadsheetExample("Change Series Arguments", SeriesActions.ChangeSeriesArgumentsAction));
            examples[6].Groups.Add(new SpreadsheetExample("Use Secondary Axes", SeriesActions.UseSecondaryAxesAction));
            examples[6].Groups.Add(new SpreadsheetExample("Remove Series", SeriesActions.RemoveSeriesAction));

            // Add nodes to the "Sparklines" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Create a Sparkline", SparklineActions.CreateSparklineGroupsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Customize Sparkline Appearance", SparklineActions.CustomizeSparklineAppearanceAction));
            examples[7].Groups.Add(new SpreadsheetExample("Rearrange Sparklines", SparklineActions.RearrangeSparklinesAction));
            examples[7].Groups.Add(new SpreadsheetExample("Specify Sparkline Axis Settings", SparklineActions.SpecifyAxisSettingsAction));

            // Add nodes to the "Style" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Set Chart Style", StyleActions.SetChartStyleAction));
            examples[8].Groups.Add(new SpreadsheetExample("Custom Series Color", StyleActions.CustomSeriesColorAction));
            examples[8].Groups.Add(new SpreadsheetExample("Set Chart Font", StyleActions.SetChartFontAction));
            examples[8].Groups.Add(new SpreadsheetExample("Set Transparency", StyleActions.TransparencyAction));

            // Add nodes to the "Titles" group of examples.
            examples[9].Groups.Add(new SpreadsheetExample("Set Title Text", TitlesActions.SetChartTitleTextAction));
            examples[9].Groups.Add(new SpreadsheetExample("Link Title to Cell Range", TitlesActions.LinkChartTitleToCellRangeAction));
            examples[9].Groups.Add(new SpreadsheetExample("Show Chart Title", TitlesActions.ShowChartTitleAction));
            examples[9].Groups.Add(new SpreadsheetExample("Set Axis Title", TitlesActions.SetAxisTitleTextAction));
            examples[9].Groups.Add(new SpreadsheetExample("Link Axis Title to Cell Range ", TitlesActions.LinkAxisTitleToCellRangeAction));
            examples[9].Groups.Add(new SpreadsheetExample("Show Axis Title", TitlesActions.ShowAxisTitleAction));


            // Add nodes to the "Trendlines" group of examples.
            examples[10].Groups.Add(new SpreadsheetExample("Display Trendline", TrendlineActions.TrendlinesAction));
            examples[10].Groups.Add(new SpreadsheetExample("Specify Trendline Label", TrendlineActions.TrendlineLabelAction));
            examples[10].Groups.Add(new SpreadsheetExample("Customize Trendline", TrendlineActions.TrendlineCustomizationAction));

            // Add nodes to the "View Options" group of examples.
            examples[11].Groups.Add(new SpreadsheetExample("Apply Gradient To a Chart Background", ViewOptionsActions.ApplyGradientToChartBackgroundAction));
            examples[11].Groups.Add(new SpreadsheetExample("Change Chart Appearance", ViewOptionsActions.ChangeChartAppearanceAction));
            examples[11].Groups.Add(new SpreadsheetExample("Custom Walls And Floor", ViewOptionsActions.CustomWallsAndFloorAction));
            examples[11].Groups.Add(new SpreadsheetExample("Specify Gap Width", ViewOptionsActions.GapWidthAction));
            examples[11].Groups.Add(new SpreadsheetExample("Show Automatic Markers", ViewOptionsActions.ShowAutomaticMarkersAction));
            examples[11].Groups.Add(new SpreadsheetExample("Show Custom Markers", ViewOptionsActions.ShowCustomMarkersAction));
            examples[11].Groups.Add(new SpreadsheetExample("Smooth Lines", ViewOptionsActions.SmoothLinesAction));
            examples[11].Groups.Add(new SpreadsheetExample("VaryColorsByPoint", ViewOptionsActions.VaryColorsByPointAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }


        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<Workbook> action = example.Action;
            action(workbook);
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile()
        {
            #region #LoadDocumentFromFile
            // Load a workbook from the file.
            workbook.LoadDocument("Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }


        private void SaveDocumentToFile()
        {
            #region #SaveDocumentToFile
            // Save the modified document to the file.
            workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
            Process.Start(new ProcessStartInfo("SavedDocument.xlsx") { UseShellExecute = true });
        }

        private void InitializeComponent()
        {
            this.treeList1 = new DevExpress.XtraTreeList.TreeList();
            this.treeListColumn1 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
            this.splitContainerControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeList1
            // 
            this.treeList1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Blue;
            this.treeList1.Appearance.FocusedCell.Options.UseFont = true;
            this.treeList1.Appearance.FocusedCell.Options.UseForeColor = true;
            this.treeList1.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.treeListColumn1});
            this.treeList1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeList1.Location = new System.Drawing.Point(0, 0);
            this.treeList1.Name = "treeList1";
            this.treeList1.OptionsBehavior.Editable = false;
            this.treeList1.OptionsView.ShowColumns = false;
            this.treeList1.OptionsView.ShowIndicator = false;
            this.treeList1.Size = new System.Drawing.Size(497, 638);
            this.treeList1.TabIndex = 0;
            // 
            // treeListColumn1
            // 
            this.treeListColumn1.Caption = "Name";
            this.treeListColumn1.FieldName = "Name";
            this.treeListColumn1.Name = "treeListColumn1";
            this.treeListColumn1.Visible = true;
            this.treeListColumn1.VisibleIndex = 0;
            this.treeListColumn1.Width = 92;
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnOpenExcel.Location = new System.Drawing.Point(0, 0);
            this.btnOpenExcel.Name = "button1";
            this.btnOpenExcel.Size = new System.Drawing.Size(497, 57);
            this.btnOpenExcel.TabIndex = 1;
            this.btnOpenExcel.Text = "Run";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // splitContainerControl1
            // 
            this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerControl1.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2;
            this.splitContainerControl1.Horizontal = false;
            this.splitContainerControl1.Location = new System.Drawing.Point(0, 0);
            this.splitContainerControl1.Name = "splitContainerControl1";
            this.splitContainerControl1.Panel1.Controls.Add(this.treeList1);
            this.splitContainerControl1.Panel1.Text = "Panel1";
            this.splitContainerControl1.Panel2.Controls.Add(this.btnOpenExcel);
            this.splitContainerControl1.Panel2.Text = "Panel2";
            this.splitContainerControl1.Size = new System.Drawing.Size(497, 700);
            this.splitContainerControl1.SplitterPosition = 57;
            this.splitContainerControl1.TabIndex = 2;
            this.splitContainerControl1.Text = "splitContainerControl1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 700);
            this.Controls.Add(this.splitContainerControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
            this.splitContainerControl1.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}

