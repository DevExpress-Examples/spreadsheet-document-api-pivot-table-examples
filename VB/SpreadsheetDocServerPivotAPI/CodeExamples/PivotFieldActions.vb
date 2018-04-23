Imports DevExpress.Spreadsheet

Namespace SpreadsheetDocServerPivotAPI
    Public NotInheritable Class PivotFieldActions

        Private Sub New()
        End Sub

        Private Shared Sub AddFieldToAxis(ByVal workbook As IWorkbook)
'            #Region "#AddToAxis"
            Dim sourceWorksheet As Worksheet = workbook.Worksheets("Data1")
            Dim worksheet As Worksheet = workbook.Worksheets.Add()
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a pivot table.
            Dim pivotTable As PivotTable = worksheet.PivotTables.Add(sourceWorksheet("A1:D41"), worksheet("B2"))

            ' Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Product"))
            ' Add the "Category" field to the column axis area.
            pivotTable.ColumnFields.Add(pivotTable.Fields("Category"))
            ' Add the "Sales" field to the data area and specify the custom field name.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Sales"), "Sales(Sum)")
            ' Specify the number format for the "Sales" field.
            dataField.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)"
            ' Add the "Region" field to the filter area.
            pivotTable.PageFields.Add(pivotTable.Fields("Region"))
'            #End Region ' #AddToAxis
        End Sub

        Private Shared Sub InsertFieldToAxis(ByVal workbook As IWorkbook)
'            #Region "#InsertAtTop"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Insert the "Region" field at the top of the row axis area.
            pivotTable.RowFields.Insert(0, pivotTable.Fields("Region"))
'            #End Region ' #InsertAtTop
        End Sub

        Private Shared Sub MoveFieldToAxis(ByVal workbook As IWorkbook)
'            #Region "#MoveToAxis"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Move the "Region" field to the column axis area.
            pivotTable.ColumnFields.Add(pivotTable.Fields("Region"))
'            #End Region ' #MoveToAxis
        End Sub

        Private Shared Sub MoveFieldUp(ByVal workbook As IWorkbook)
'            #Region "#MoveUp"
            Dim worksheet As Worksheet = workbook.Worksheets("Report3")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Move the "Category" field one position up in the row area.
            pivotTable.RowFields("Category").MoveUp()
'            #End Region ' #MoveUp
        End Sub

        Private Shared Sub MoveFieldDown(ByVal workbook As IWorkbook)
'            #Region "#MoveDown"
            Dim worksheet As Worksheet = workbook.Worksheets("Report3")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Move the "Region" field one position down in the row area.
            pivotTable.RowFields("Region").MoveDown()
'            #End Region ' #MoveDown
        End Sub

        Private Shared Sub RemoveFieldFromAxis(ByVal workbook As IWorkbook)
'            #Region "#RemoveFromAxis"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Remove the "Product" field from the row axis area.
            pivotTable.RowFields.Remove(pivotTable.RowFields("Product"))
'            #End Region ' #RemoveFromAxis
        End Sub

        Private Shared Sub SortFieldItems(ByVal workbook As IWorkbook)
'            #Region "#SortFieldItems"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the pivot field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("Product")
            ' Sort items in the "Product" field. 
            field.SortType = PivotFieldSortType.Ascending
'            #End Region ' #SortFieldItems
        End Sub

        Private Shared Sub SortFieldItemsByDataField(ByVal workbook As IWorkbook)
'            #Region "#SortFieldItemsByDataField"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the pivot field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("Product")
            ' Sort items in the "Product" field by values of the "Sum of Sales" data field.
            field.SortItems(PivotFieldSortType.Ascending, 0)
'            #End Region ' #SortFieldItemsByDataField
        End Sub

        Private Shared Sub MultipleSubtotals(ByVal workbook As IWorkbook)
'            #Region "#MultipleSubtotals"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the pivot field by its name in the collection.
            Dim field As PivotField = pivotTable.Fields("Category")
            ' Display multiple subtotals for the field.  
            field.SetSubtotal(PivotSubtotalFunctions.Sum Or PivotSubtotalFunctions.Average)
'            #End Region ' #MultipleSubtotals
        End Sub
    End Class
End Namespace

