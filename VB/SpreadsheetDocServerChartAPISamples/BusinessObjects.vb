Imports DevExpress.Spreadsheet
Imports DevExpress.XtraTreeList
Imports System
Imports System.ComponentModel

Namespace SpreadsheetDocServerChartAPISamples
	Public Class SpreadsheetNode
		Private _groups As New GroupsOfSpreadsheetExamples()
		Private _owner As GroupsOfSpreadsheetExamples

		Public Sub New(ByVal name As String)
			Me.Name = name
		End Sub
		<Browsable(False)>
		Public ReadOnly Property Groups() As GroupsOfSpreadsheetExamples
			Get
				Return _groups
			End Get
		End Property
		Public Property Name() As String

		<Browsable(False)>
		Public Property Owner() As GroupsOfSpreadsheetExamples
			Get
				Return _owner
			End Get
			Set(ByVal value As GroupsOfSpreadsheetExamples)
				_owner = value
			End Set
		End Property
	End Class

	Public Class SpreadsheetExample
		Inherits SpreadsheetNode

		Public Sub New(ByVal name As String, ByVal action As Action(Of Workbook))
			MyBase.New(name)
			Me.Action = action
		End Sub
		Private privateAction As Action(Of Workbook)
		Public Property Action() As Action(Of Workbook)
			Get
				Return privateAction
			End Get
			Private Set(ByVal value As Action(Of Workbook))
				privateAction = value
			End Set
		End Property
	End Class

	Public Class GroupsOfSpreadsheetExamples
		Inherits BindingList(Of SpreadsheetNode)
		Implements TreeList.IVirtualTreeListData

		Private Sub IVirtualTreeListData_VirtualTreeGetChildNodes(ByVal info As VirtualTreeGetChildNodesInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetChildNodes
			Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
			info.Children = obj.Groups
		End Sub
		Protected Overrides Sub InsertItem(ByVal index As Integer, ByVal item As SpreadsheetNode)
			item.Owner = Me
			MyBase.InsertItem(index, item)
		End Sub
		Private Sub IVirtualTreeListData_VirtualTreeGetCellValue(ByVal info As VirtualTreeGetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetCellValue
			Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
			Select Case info.Column.Caption
				Case "Name"
					info.CellData = obj.Name
			End Select
		End Sub
		Private Sub IVirtualTreeListData_VirtualTreeSetCellValue(ByVal info As VirtualTreeSetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeSetCellValue
			Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
			Select Case info.Column.Caption
				Case "Name"
					obj.Name = CStr(info.NewCellData)
			End Select
		End Sub
	End Class
End Namespace
