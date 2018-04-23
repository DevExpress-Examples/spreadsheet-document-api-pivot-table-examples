Imports DevExpress.Spreadsheet
Imports System.Collections.Generic

Namespace SpreadsheetDocServerPivotAPI
    Friend Class PivotFieldGroupingActions
        Private Shared Sub GroupFieldItems(ByVal workbook As IWorkbook)
'            #Region "#GroupFieldItems"
            Dim worksheet As Worksheet = workbook.Worksheets("Report11")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "State" field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("State")
            ' Add the "State" field to the column axis area.
            pivotTable.ColumnFields.Add(field)

            ' Group the first three items in the field.
            Dim items As IEnumerable(Of Integer) = New List(Of Integer)() From {0, 1, 2}
            field.GroupItems(items)
            ' Access the created grouped field by its index in the field collection.
            Dim groupedFieldIndex As Integer = pivotTable.Fields.Count - 1
            Dim groupedField As PivotField = pivotTable.Fields(groupedFieldIndex)
            ' Set the grouped item caption to "West".
            groupedField.Items(0).Caption = "West"
'            #End Region ' #GroupFieldItems
        End Sub

        Private Shared Sub GroupFieldByDates(ByVal workbook As IWorkbook)
'            #Region "#GroupFieldByDates"
            Dim worksheet As Worksheet = workbook.Worksheets("Report8")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "DATE" field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("DATE")
            ' Group field items by quarters and months.
            field.GroupItems(PivotFieldGroupByType.Quarters Or PivotFieldGroupByType.Months)
'            #End Region ' #GroupFieldByDates
        End Sub

        Private Shared Sub GroupFieldByNumericRanges(ByVal workbook As IWorkbook)
'            #Region "#GroupFieldByNumericRanges"
            Dim worksheet As Worksheet = workbook.Worksheets("Report12")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "Usual Hours Worked" field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("Sales")
            ' Group field items from 1000 to 4000 by 1000.
            field.GroupItems(1000, 4000, 1000, PivotFieldGroupByType.NumericRanges)
'            #End Region ' #GroupFieldByNumericRanges
        End Sub

        Private Shared Sub UngroupSpecificItem(ByVal workbook As IWorkbook)
'            #Region "#UngroupSpecificItem"
            Dim worksheet As Worksheet = workbook.Worksheets("Report11")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "State" field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("State")
            ' Add the "State" field to the column axis area.
            pivotTable.ColumnFields.Add(field)

            ' Group the first three items in the field.
            Dim items As IEnumerable(Of Integer) = New List(Of Integer)() From {0, 1, 2}
            field.GroupItems(items)
            ' Access the created grouped field by its index in the field collection.
            Dim groupedFieldIndex As Integer = pivotTable.Fields.Count - 1
            Dim groupedField As PivotField = pivotTable.Fields(groupedFieldIndex)
            ' Set the grouped item caption to "West".
            groupedField.Items(0).Caption = "West"

            ' Group the remaining field items.
            items = New List(Of Integer)() From {3, 4, 5}
            field.GroupItems(items)
            ' Set the grouped item caption to "Midwest"
            groupedField.Items(1).Caption = "Midwest"

            ' Ungroup the "West" item.
            items = New List(Of Integer) From {0}
            groupedField.UngroupItems(items)
'            #End Region ' #UngroupSpecificItem
        End Sub

        Private Shared Sub UngroupFieldItems(ByVal workbook As IWorkbook)
'            #Region "#UngroupFieldItems"
            Dim worksheet As Worksheet = workbook.Worksheets("Report8")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "DATE" field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("DATE")
            ' Group field items by days.
            field.GroupItems(field.GroupingInfo.DefaultStartValue, field.GroupingInfo.DefaultEndValue, 50, PivotFieldGroupByType.Days)
            ' Ungroup field items.
            field.UngroupItems()
'            #End Region ' #UngroupFieldItems
        End Sub
    End Class
End Namespace
