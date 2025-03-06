Imports DevExpress.Spreadsheet
Imports System

Namespace SpreadsheetDocServerPivotAPI
	Public Module PivotTableLayoutActions
		Public ColumnGrandTotalsAction As Action(Of Workbook) = AddressOf ColumnGrandTotals
		Public RowGrandTotalsAction As Action(Of Workbook) = AddressOf RowGrandTotals
		Public DataOnRowsAction As Action(Of Workbook) = AddressOf DataOnRows
		Public MergeTitlesAction As Action(Of Workbook) = AddressOf MergeTitles
		Public ShowAllSubtotalsAction As Action(Of Workbook) = AddressOf ShowAllSubtotals
		Public HideAllSubtotalsAction As Action(Of Workbook) = AddressOf HideAllSubtotals
		Public SetCompactReportLayoutAction As Action(Of Workbook) = AddressOf SetCompactReportLayout
		Public SetOutlineReportLayoutAction As Action(Of Workbook) = AddressOf SetOutlineReportLayout
		Public SetTabularReportLayoutAction As Action(Of Workbook) = AddressOf SetTabularReportLayout
		Public RepeatAllItemLabelsAction As Action(Of Workbook) = AddressOf RepeatAllItemLabels
		Public InsertBlankRowsAction As Action(Of Workbook) = AddressOf InsertBlankRows
		Public RemoveBlankRowsAction As Action(Of Workbook) = AddressOf RemoveBlankRows

		Private Sub ColumnGrandTotals(ByVal workbook As IWorkbook)
#Region "#ColumnGrandTotals"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Add the "Region" field to the column axis area. 
			pivotTable.ColumnFields.Add(pivotTable.Fields("Region"))

			' Hide grand totals for columns.
			pivotTable.Layout.ShowColumnGrandTotals = False
#End Region ' #ColumnGrandTotals
		End Sub

		Private Sub RowGrandTotals(ByVal workbook As IWorkbook)
#Region "#RowGrandTotals"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Add the "Region" field to the column axis area. 
			pivotTable.ColumnFields.Add(pivotTable.Fields("Region"))

			' Hide grand totals for rows.
			pivotTable.Layout.ShowRowGrandTotals = False
#End Region ' #RowGrandTotals
		End Sub

		Private Sub DataOnRows(ByVal workbook As IWorkbook)
#Region "#MultipleDataFields"
			Dim worksheet As Worksheet = workbook.Worksheets("Report2")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display value fields in separate columns.
			pivotTable.Layout.DataOnRows = False
#End Region ' #MultipleDataFields
		End Sub
		Private Sub MergeTitles(ByVal workbook As IWorkbook)
#Region "#MergeTitles"
			Dim worksheet As Worksheet = workbook.Worksheets("Report4")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display the pivot table in the tabular form.
			pivotTable.Layout.SetReportLayout(PivotReportLayout.Tabular)
			' Merge and center cells with labels. 
			pivotTable.Layout.MergeTitles = True
#End Region ' #MergeTitles
		End Sub

		Private Sub ShowAllSubtotals(ByVal workbook As IWorkbook)
#Region "#ShowAllSubtotals"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Show all subtotals at the top of each group.
			pivotTable.Layout.ShowAllSubtotals(True)
#End Region ' #ShowAllSubtotals
		End Sub

		Private Sub HideAllSubtotals(ByVal workbook As IWorkbook)
#Region "#HideAllSubtotals"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Hide subtotals at the top of each group.
			pivotTable.Layout.HideAllSubtotals()
#End Region ' #HideAllSubtotals
		End Sub

		Private Sub SetCompactReportLayout(ByVal workbook As IWorkbook)
#Region "#CompactReportLayout"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display the pivot table in the compact form.
			pivotTable.Layout.SetReportLayout(PivotReportLayout.Compact)
#End Region ' #CompactReportLayout
		End Sub

		Private Sub SetOutlineReportLayout(ByVal workbook As IWorkbook)
#Region "#OutlineReportLayout"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display the pivot table in the outline form.
			pivotTable.Layout.SetReportLayout(PivotReportLayout.Outline)
#End Region ' #OutlineReportLayout
		End Sub

		Private Sub SetTabularReportLayout(ByVal workbook As IWorkbook)
#Region "#TabularReportLayout"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display the pivot table in the tabular form.
			pivotTable.Layout.SetReportLayout(PivotReportLayout.Tabular)
#End Region ' #TabularReportLayout
		End Sub

		Private Sub RepeatAllItemLabels(ByVal workbook As IWorkbook)
#Region "#RepeatAllItemLabels"
			Dim worksheet As Worksheet = workbook.Worksheets("Report5")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Display repeated column labels.
			pivotTable.Layout.RepeatAllItemLabels(True)
#End Region ' #RepeatAllItemLabels
		End Sub

		Private Sub InsertBlankRows(ByVal workbook As IWorkbook)
#Region "#InsertBlankRows"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Insert a blank row after each group of items.
			pivotTable.Layout.InsertBlankRows()
#End Region ' #InsertBlankRows
		End Sub

		Private Sub RemoveBlankRows(ByVal workbook As IWorkbook)
#Region "#RemoveBlankRows"
			Dim worksheet As Worksheet = workbook.Worksheets("Report1")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
			' Insert a blank row after each group of items.
			pivotTable.Layout.InsertBlankRows()

			' Remove a blank row after each group of items.
			pivotTable.Layout.RemoveBlankRows()
#End Region ' #RemoveBlankRows
		End Sub
	End Module
End Namespace
