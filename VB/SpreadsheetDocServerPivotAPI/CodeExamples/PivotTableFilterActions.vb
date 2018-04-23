Imports DevExpress.Spreadsheet

Namespace SpreadsheetDocServerPivotAPI
    Public NotInheritable Class PivotTableFilterActions

        Private Sub New()
        End Sub

        Private Shared Sub SetItemFilter(ByVal workbook As IWorkbook)
'            #Region "#ItemFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Show the first item in the "Product" field.
            pivotTable.Fields(1).ShowSingleItem(0)

            'Show all items in the "Product" field (the default option).
            'pivotTable.Fields[1].ShowAllItems();
'            #End Region ' #ItemFilter
        End Sub

        Private Shared Sub SetItemVisibilityFilter(ByVal workbook As IWorkbook)
'            #Region "#ItemVisibility"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access items of the "Product" field.
            Dim pivotFieldItems As PivotItemCollection = pivotTable.Fields(1).Items

            ' Hide the first item in the "Product" field.
            pivotFieldItems(0).Visible = False
'            #End Region ' #ItemVisibility
        End Sub

        Private Shared Sub SetLabelFilter(ByVal workbook As IWorkbook)
'            #Region "#LabelFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "Region" field.
            Dim field As PivotField = pivotTable.Fields(0)
            ' Filter the "Region" field by text to display sales data for the "South" region.  
            pivotTable.Filters.Add(field, PivotFilterType.CaptionEqual, "South")
'            #End Region ' #LabelFilter
        End Sub

        Private Shared Sub SetValueFilter(ByVal workbook As IWorkbook)
'            #Region "#ValueFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "Product" field.
            Dim field As PivotField = pivotTable.Fields(1)
            ' Filter the "Product" field to display products with grand total sales between $6000 and $13000.
            pivotTable.Filters.Add(field, pivotTable.DataFields(0), PivotFilterType.ValueBetween, 6000, 13000)
'            #End Region ' #ValueFilter
        End Sub

        Private Shared Sub SetTop10Filter(ByVal workbook As IWorkbook)
'            #Region "#Top10Filter"
            Dim worksheet As Worksheet = workbook.Worksheets("Report4")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "Product" field.
            Dim field As PivotField = pivotTable.Fields(1)
            ' Filter the "Product" field to display two products with the lowest sales.
            Dim filter As PivotFilter = pivotTable.Filters.Add(field, pivotTable.DataFields(0), PivotFilterType.Count, 2)
            filter.Top10Type = PivotFilterTop10Type.Bottom
'            #End Region ' #Top10Filter
        End Sub

        Private Shared Sub SetDateFilter(ByVal workbook As IWorkbook)
'            #Region "#DateFilter"
            Dim worksheet As Worksheet = workbook.Worksheets("Report6")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")
            ' Access the "Date" field.
            Dim field As PivotField = pivotTable.Fields(0)
            ' Filter the "Date" field to display sales for the second quarter.
            pivotTable.Filters.Add(field, PivotFilterType.SecondQuarter)
'            #End Region ' #DateFilter
        End Sub

        Private Shared Sub SetMultipleFilter(ByVal workbook As IWorkbook)
'            #Region "#MultipleFilters"
            Dim worksheet As Worksheet = workbook.Worksheets("Report6")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection.
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Allow multiple filters for a field.
            pivotTable.Behavior.AllowMultipleFieldFilters = True

            ' Filter the "Date" field to display sales for the second quarter.
            Dim field1 As PivotField = pivotTable.Fields(0)
            pivotTable.Filters.Add(field1, PivotFilterType.SecondQuarter)

            ' Add the second filter to the "Date" field to display two days with the lowest sales.  
            Dim filter As PivotFilter = pivotTable.Filters.Add(field1, pivotTable.DataFields(0), PivotFilterType.Count, 2)
            filter.Top10Type = PivotFilterTop10Type.Bottom
'            #End Region ' #MultipleFilters
        End Sub
    End Class
End Namespace

