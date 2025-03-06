Imports DevExpress.Spreadsheet
Imports DevExpress.XtraEditors
Imports DevExpress.XtraTab
Imports DevExpress.XtraTreeList
Imports DevExpress.XtraTreeList.Columns
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms

Namespace SpreadsheetDocServerPivotAPI
	Partial Public Class Form1
		Inherits DevExpress.XtraEditors.XtraForm

		Private workbook As New Workbook()
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
			examples.Add(New SpreadsheetNode("Pivot Calculated Fields"))
			examples.Add(New SpreadsheetNode("Pivot Calculated Item"))
			examples.Add(New SpreadsheetNode("Pivot Fields"))
			examples.Add(New SpreadsheetNode("Pivot Field Groups"))
			examples.Add(New SpreadsheetNode("Pivot Tables"))
			examples.Add(New SpreadsheetNode("Pivot Table Filter"))
			examples.Add(New SpreadsheetNode("Pivot Table Formatting"))
			examples.Add(New SpreadsheetNode("Pivot Table Layout"))
			examples.Add(New SpreadsheetNode("Value Field Settings"))

#End Region

#Region "ExampleNodes"
			' Add nodes to the "Calculated Field" group of examples.
			examples(0).Groups.Add(New SpreadsheetExample("Add Calculated Field", PivotCalculatedFieldActions.AddCalculatedFieldAction))
			examples(0).Groups.Add(New SpreadsheetExample("Modify Calculated Field", PivotCalculatedFieldActions.ModifyCalculatedFieldAction))
			examples(0).Groups.Add(New SpreadsheetExample("Remove Calculated Field", PivotCalculatedFieldActions.RemoveCalculatedFieldAction))

			' Add nodes to the "Calculated Item" group of examples.
			examples(1).Groups.Add(New SpreadsheetExample("Add Calculated Item", PivotCalculatedItemActions.AddCalculatedItemAction))
			examples(1).Groups.Add(New SpreadsheetExample("Modify Calculated Item", PivotCalculatedItemActions.ModifyCalculatedItemAction))
			examples(1).Groups.Add(New SpreadsheetExample("Remove Calculated Item", PivotCalculatedItemActions.RemoveCalculatedItemAction))

			' Add nodes to the "Pivot Field" group of examples.
			examples(2).Groups.Add(New SpreadsheetExample("Add Field to Axis", PivotFieldActions.AddFieldToAxisAction))
			examples(2).Groups.Add(New SpreadsheetExample("Insert Field to Axis", PivotFieldActions.InsertFieldToAxisAction))
			examples(2).Groups.Add(New SpreadsheetExample("Move Field Down", PivotFieldActions.MoveFieldDownAction))
			examples(2).Groups.Add(New SpreadsheetExample("Move Field to Axis", PivotFieldActions.MoveFieldToAxisAction))
			examples(2).Groups.Add(New SpreadsheetExample("Move Field Up", PivotFieldActions.MoveFieldUpAction))
			examples(2).Groups.Add(New SpreadsheetExample("Multiple Subtotals", PivotFieldActions.MultipleSubtotalsAction))
			examples(2).Groups.Add(New SpreadsheetExample("Remove Field from Axis", PivotFieldActions.RemoveFieldFromAxisAction))
			examples(2).Groups.Add(New SpreadsheetExample("Sort Field Items", PivotFieldActions.SortFieldItemsAction))
			examples(2).Groups.Add(New SpreadsheetExample("Sort Field Items by Data Field", PivotFieldActions.SortFieldItemsByDataFieldAction))


			' Add nodes to the "Chart Legend" group of examples.
			examples(3).Groups.Add(New SpreadsheetExample("Group Field by Dates", PivotFieldGroupingActions.GroupFieldByDatesAction))
			examples(3).Groups.Add(New SpreadsheetExample("Group Field Items", PivotFieldGroupingActions.GroupFieldItemsAction))
			examples(3).Groups.Add(New SpreadsheetExample("Ungroup Field Items", PivotFieldGroupingActions.UngroupFieldItemsAction))
			examples(3).Groups.Add(New SpreadsheetExample("Ungroup Specific Item", PivotFieldGroupingActions.UngroupSpecificItemAction))

			' Add nodes to the "Pivot Table" group of examples.        
			examples(4).Groups.Add(New SpreadsheetExample("Create Pivot Table from Cache", PivotTableActions.CreatePivotTableFromCacheAction))
			examples(4).Groups.Add(New SpreadsheetExample("Create Pivot Table from Range", PivotTableActions.CreatePivotTableFromRangeAction))
			examples(4).Groups.Add(New SpreadsheetExample("Change Behavior Options", PivotTableActions.ChangeBehaviorOptionsAction))
			examples(4).Groups.Add(New SpreadsheetExample("Change Pivot Table Data Source", PivotTableActions.ChangePivotTableDataSourceAction))
			examples(4).Groups.Add(New SpreadsheetExample("Change Pivot Table Location", PivotTableActions.ChangePivotTableLocationAction))
			examples(4).Groups.Add(New SpreadsheetExample("Clear Pivot Table", PivotTableActions.ClearPivotTableAction))
			examples(4).Groups.Add(New SpreadsheetExample("Move Pivot Table to Worksheet", PivotTableActions.MovePivotTableToWorksheetAction))
			examples(4).Groups.Add(New SpreadsheetExample("Remove Pivot Table", PivotTableActions.RemovePivotTableAction))

			' Add nodes to the "Pivot Table Filter" group of examples.
			examples(5).Groups.Add(New SpreadsheetExample("Set Item Filter", PivotTableFilterActions.SetItemFilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Label Filter", PivotTableFilterActions.SetLabelFilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Date Filter", PivotTableFilterActions.SetDateFilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Multiple Filter", PivotTableFilterActions.SetMultipleFilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Top 10 Filter", PivotTableFilterActions.SetTop10FilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Value Filter", PivotTableFilterActions.SetValueFilterAction))
			examples(5).Groups.Add(New SpreadsheetExample("Set Item Visibility Filter", PivotTableFilterActions.SetItemVisibilityFilterAction))

			' Add nodes to the "Pivot Table Formatting" group of examples.
			examples(6).Groups.Add(New SpreadsheetExample("Banded Columns", PivotTableFormattingActions.BandedColumnsAction))
			examples(6).Groups.Add(New SpreadsheetExample("Banded Rows", PivotTableFormattingActions.BandedRowsAction))
			examples(6).Groups.Add(New SpreadsheetExample("Change Pivot Table Style", PivotTableFormattingActions.ChangeStylePivotTableAction))
			examples(6).Groups.Add(New SpreadsheetExample("Show Column Headers", PivotTableFormattingActions.ShowColumnHeadersAction))
			examples(6).Groups.Add(New SpreadsheetExample("Show Row Headers", PivotTableFormattingActions.ShowRowHeadersAction))

			' Add nodes to the "Pivot Table Layout" group of examples.
			examples(7).Groups.Add(New SpreadsheetExample("Column Grand Totals", PivotTableLayoutActions.ColumnGrandTotalsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Data On Rows", PivotTableLayoutActions.DataOnRowsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Hide All Subtotals", PivotTableLayoutActions.HideAllSubtotalsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Insert Blank Rows", PivotTableLayoutActions.InsertBlankRowsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Merge Titles", PivotTableLayoutActions.MergeTitlesAction))
			examples(7).Groups.Add(New SpreadsheetExample("Remove Blank Rows", PivotTableLayoutActions.RemoveBlankRowsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Repeat All Item Labels", PivotTableLayoutActions.RepeatAllItemLabelsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Row Grand Totals", PivotTableLayoutActions.RowGrandTotalsAction))
			examples(7).Groups.Add(New SpreadsheetExample("Set Compact Report Layout", PivotTableLayoutActions.SetCompactReportLayoutAction))
			examples(7).Groups.Add(New SpreadsheetExample("Set Outline Report Layout", PivotTableLayoutActions.SetOutlineReportLayoutAction))
			examples(7).Groups.Add(New SpreadsheetExample("Set Tabular Report Layout", PivotTableLayoutActions.SetTabularReportLayoutAction))
			examples(7).Groups.Add(New SpreadsheetExample("Show All Subtotals", PivotTableLayoutActions.ShowAllSubtotalsAction))

			' Add nodes to the "Value Field Settings" group of examples.
			examples(8).Groups.Add(New SpreadsheetExample("Change Summary Function", ValueFieldSettingsActions.ChangeSummaryFunctionAction))
			examples(8).Groups.Add(New SpreadsheetExample("Difference From", ValueFieldSettingsActions.DifferenceFromAction))
			examples(8).Groups.Add(New SpreadsheetExample("Number Format", ValueFieldSettingsActions.NumberFormatAction))
			examples(8).Groups.Add(New SpreadsheetExample("Percent Of", ValueFieldSettingsActions.PercentOfAction))
			examples(8).Groups.Add(New SpreadsheetExample("Percent Of Parent Row", ValueFieldSettingsActions.PercentOfParentRowTotalAction))
			examples(8).Groups.Add(New SpreadsheetExample("Rank Largest to Smallest", ValueFieldSettingsActions.RankLargestToSmallestAction))
			examples(8).Groups.Add(New SpreadsheetExample("Running Total In", ValueFieldSettingsActions.RunningTotalInAction))
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
			workbook.LoadDocument("PivotTableTemplate.xlsx", DocumentFormat.OpenXml)
#End Region ' #LoadDocumentFromFile
		End Sub


		Private Sub SaveDocumentToFile()
#Region "#SaveDocumentToFile"
			' Save the modified document to the file.
			workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml)
#End Region ' #SaveDocumentToFile
			Process.Start(New ProcessStartInfo("SavedDocument.xlsx") With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
