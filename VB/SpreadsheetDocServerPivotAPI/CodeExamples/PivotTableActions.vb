Imports DevExpress.Spreadsheet

Namespace SpreadsheetDocServerPivotAPI
    Public NotInheritable Class PivotTableActions

        Private Sub New()
        End Sub


        Private Shared Sub CreatePivotTableFromRange(ByVal workbook As IWorkbook)
'            #Region "#CreateFromRange"
            Dim sourceWorksheet As Worksheet = workbook.Worksheets("Data1")
            Dim worksheet As Worksheet = workbook.Worksheets.Add()
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a pivot table using the cell range "A1:D41" as the data source.
            Dim pivotTable As PivotTable = worksheet.PivotTables.Add(sourceWorksheet("A1:D41"), worksheet("B2"))

            ' Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Category"))
            ' Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Product"))
            ' Add the "Sales" field to the data area.
            pivotTable.DataFields.Add(pivotTable.Fields("Sales"))
'            #End Region ' #CreateFromRange
        End Sub

        Private Shared Sub CreatePivotTableFromCache(ByVal workbook As IWorkbook)
'            #Region "#CreateFromPivotCache"
            Dim worksheet As Worksheet = workbook.Worksheets.Add()
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a pivot table based on the specified PivotTable cache.
            Dim cache As PivotCache = workbook.Worksheets("Report1").PivotTables("PivotTable1").Cache
            Dim pivotTable As PivotTable = worksheet.PivotTables.Add(cache, worksheet("B2"))

            ' Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Category"))
            ' Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Product"))
            ' Add the "Sales" field to the data area.
            pivotTable.DataFields.Add(pivotTable.Fields("Sales"))

            ' Set the default style for the pivot table.
            pivotTable.Style = workbook.TableStyles.DefaultPivotStyle

'            #End Region ' #CreateFromPivotCache
        End Sub

        Private Shared Sub RemovePivotTable(ByVal workbook As IWorkbook)
'            #Region "#RemoveTable"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Remove the pivot table from the collection.
            worksheet.PivotTables.Remove(pivotTable)

'            #End Region ' #RemoveTable
        End Sub
        Private Shared Sub ChangePivotTableLocation(ByVal workbook As IWorkbook)
'            #Region "#ChangeLocation"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Change the pivot table location.
            worksheet.PivotTables("PivotTable1").MoveTo(worksheet("A7"))
            ' Refresh the pivot table.
            worksheet.PivotTables("PivotTable1").Cache.Refresh()

'            #End Region ' #ChangeLocation
        End Sub
        Private Shared Sub MovePivotTableToWorksheet(ByVal workbook As IWorkbook)
'            #Region "#MoveToWorksheet"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")

            ' Create a new worksheet.
            Dim targetWorksheet As Worksheet = workbook.Worksheets.Add()

            ' Access the pivot table by its name in the collection
            ' and move it to the new worksheet.
            worksheet.PivotTables("PivotTable1").MoveTo(targetWorksheet("B2"))
            ' Refresh the pivot table.
            targetWorksheet.PivotTables("PivotTable1").Cache.Refresh()

            workbook.Worksheets.ActiveWorksheet = targetWorksheet
'            #End Region ' #MoveToWorksheet
        End Sub

        Private Shared Sub ChangePivotTableDataSource(ByVal workbook As IWorkbook)
'            #Region "#ChangeDataSource"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            Dim sourceWorksheet As Worksheet = workbook.Worksheets("Data2")
            ' Change the data source of the pivot table.
            pivotTable.ChangeDataSource(sourceWorksheet("A1:H6367"))

            ' Add the "State" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("State"))
            ' Add the "Yearly Earnings" field to the data area.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Yearly Earnings"))
            ' Calculate the average of the "Yearly Earnings" values for each state.
            dataField.SummarizeValuesBy = PivotDataConsolidationFunction.Average

'            #End Region ' #ChangeDataSource
        End Sub

        Private Shared Sub ClearPivotTable(ByVal workbook As IWorkbook)
'            #Region "#ClearTable"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Clear the pivot table.
            worksheet.PivotTables("PivotTable1").Clear()
'            #End Region ' #ClearTable
        End Sub

        Private Shared Sub ChangeBehaviorOptions(ByVal workbook As IWorkbook)
'            #Region "#ChangeBehaviorOptions"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet
            worksheet.Columns("B").WidthInCharacters = 40

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Restrict specific operations for the pivot table.
            Dim behaviorOptions As PivotBehaviorOptions = pivotTable.Behavior
            behaviorOptions.AutoFitColumns = False
            behaviorOptions.EnableFieldList = False

            ' Refresh the pivot table.
            pivotTable.Cache.Refresh()
'            #End Region ' #ChangeBehaviorOptions
        End Sub
    End Class
End Namespace
