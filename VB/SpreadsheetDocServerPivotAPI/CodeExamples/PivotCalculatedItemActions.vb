Imports DevExpress.Charts.Model
Imports DevExpress.Spreadsheet
Imports System

Namespace SpreadsheetDocServerPivotAPI
	Friend Class PivotCalculatedItemActions

		Public Shared AddCalculatedItemAction As Action(Of Workbook) = AddressOf AddCalculatedItem
		Public Shared RemoveCalculatedItemAction As Action(Of Workbook) = AddressOf RemoveCalculatedItem
		Public Shared ModifyCalculatedItemAction As Action(Of Workbook) = AddressOf ModifyCalculatedItem
		Private Shared Sub AddCalculatedItem(ByVal workbook As IWorkbook)
#Region "#AddCalculatedItem"
			Dim worksheet As Worksheet = workbook.Worksheets("Report10")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

			' Access the pivot field by its name in the collection.
			Dim field As PivotField = pivotTable.Fields("State")

			' Add calculated items to the "State" field.
			field.CalculatedItems.Add("=Arizona+California+Colorado", "West Total")
			field.CalculatedItems.Add("=Illinois+Kansas+Wisconsin", "Midwest Total")
#End Region ' #AddCalculatedItem
		End Sub

		Private Shared Sub RemoveCalculatedItem(ByVal workbook As IWorkbook)
#Region "#RemoveCalculatedItem"
			Dim worksheet As Worksheet = workbook.Worksheets("Report7")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

			' Access the pivot field by its name in the collection.
			Dim field As PivotField = pivotTable.Fields("Customer")

			' Add a calculated item to the "Customer" field.
			field.CalculatedItems.Add("='Big Foods'*110%", "Big Foods Sales Plan")

			'Remove the calculated item by its index from the collection.
			field.CalculatedItems.RemoveAt(0)
#End Region ' #RemoveCalculatedItem
		End Sub

		Private Shared Sub ModifyCalculatedItem(ByVal workbook As IWorkbook)
#Region "#ModifyCalculatedItem"
			Dim worksheet As Worksheet = workbook.Worksheets("Report7")
			workbook.Worksheets.ActiveWorksheet = worksheet

			' Access the pivot table by its name in the collection.
			Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

			' Access the pivot field by its name in the collection.
			Dim field As PivotField = pivotTable.Fields("Customer")

			' Add a calculated item to the "Customer" field.
			Dim item As PivotItem = field.CalculatedItems.Add("='Big Foods'*110%", "Big Foods Sales Plan")

			'Change the formula for the calculated item.
			item.Formula = "='Big Foods'*115%"
#End Region ' #ModifyCalculatedItem
		End Sub
	End Class
End Namespace
