Imports DevExpress.Spreadsheet

Namespace SpreadsheetDocServerPivotAPI
    Public NotInheritable Class PivotTableFormattingActions

        Private Sub New()
        End Sub


        Private Shared Sub ChangeStylePivotTable(ByVal workbook As IWorkbook)
'            #Region "#SetStyle"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Set the pivot table style.
            pivotTable.Style = workbook.TableStyles(BuiltInPivotStyleId.PivotStyleDark7)
'            #End Region ' #SetStyle
        End Sub


        Private Shared Sub BandedColumns(ByVal workbook As IWorkbook)
'            #Region "#BandedColumns"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Apply the banded column formatting to the pivot table. 
            pivotTable.BandedColumns = True
'            #End Region ' #BandedColumns
        End Sub

        Private Shared Sub BandedRows(ByVal workbook As IWorkbook)
'            #Region "#BandedRows"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Apply the banded row formatting to the pivot table. 
            pivotTable.BandedRows = True
'            #End Region ' #BandedRows
        End Sub

        Private Shared Sub ShowColumnHeaders(ByVal workbook As IWorkbook)
'            #Region "#ColumnHeaders"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Remove formatting from column headers.
            pivotTable.ShowColumnHeaders = False
'            #End Region ' #ColumnHeaders
        End Sub

        Private Shared Sub ShowRowHeaders(ByVal workbook As IWorkbook)
'            #Region "#RowHeaders"
            Dim worksheet As Worksheet = workbook.Worksheets("Report1")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Add the "Region" field to the column axis area. 
            pivotTable.ColumnFields.Add(pivotTable.Fields("Region"))
            ' Remove formatting from row headers.
            pivotTable.ShowRowHeaders = False
'            #End Region ' #RowHeaders
        End Sub
    End Class
End Namespace
