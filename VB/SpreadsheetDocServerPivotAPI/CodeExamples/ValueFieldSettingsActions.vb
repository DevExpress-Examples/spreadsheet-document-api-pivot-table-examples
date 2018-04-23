Imports DevExpress.Spreadsheet

Namespace SpreadsheetDocServerPivotAPI
    Public NotInheritable Class ValueFieldSettingsActions

        Private Sub New()
        End Sub

        Private Shared Sub ChangeSummaryFunction(ByVal workbook As IWorkbook)
'            #Region "#ChangeSummaryFunction"
            Dim sourceWorksheet As Worksheet = workbook.Worksheets("Data5")
            Dim worksheet As Worksheet = workbook.Worksheets.Add()
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a pivot table using the cell range "A1:E65" as the data source.
            Dim pivotTable As PivotTable = worksheet.PivotTables.Add(sourceWorksheet("A1:E65"), worksheet("B2"))

            ' Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Category"))
            ' Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Product"))

            ' Add the "Amount" field to the data area.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Amount"))
            ' Use the "Average" function to summarize values in the data field.
            dataField.SummarizeValuesBy = PivotDataConsolidationFunction.Average
            ' Specify the number format for the data field.
            dataField.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)"
'            #End Region ' #ChangeSummaryFunction
        End Sub

        Private Shared Sub DifferenceFrom(ByVal workbook As IWorkbook)
'            #Region "#DifferenceFrom"
            Dim worksheet As Worksheet = workbook.Worksheets("Report14")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Access the data field by its index in the collection.
            Dim dataField As PivotDataField = pivotTable.DataFields(0)
            ' Display the difference in product sales between the current quarter and the previous quarter.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.Difference, pivotTable.Fields("Quarter"), PivotBaseItemType.Previous)
'            #End Region ' #DifferenceFrom
        End Sub

        Private Shared Sub PercentOf(ByVal workbook As IWorkbook)
'            #Region "#PercentOf"
            Dim worksheet As Worksheet = workbook.Worksheets("Report14")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Access the data field by its index in the collection.
            Dim dataField As PivotDataField = pivotTable.DataFields(0)
            ' Select the base field ("Quarter"). 
            Dim baseField As PivotField = pivotTable.Fields("Quarter")
            ' Select the base item ("Q1"). 
            Dim baseItem As PivotItem = baseField.Items(0)
            ' Show values as the percentage of the value of the base item in the base field. 
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.Percent, baseField, baseItem)
'            #End Region ' #PercentOf
        End Sub

        Private Shared Sub PercentOfParentRowTotal(ByVal workbook As IWorkbook)
'            #Region "#PercentOfParentRowTotal"
            Dim worksheet As Worksheet = workbook.Worksheets("Report16")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Amount"), "% of Parent Row Total")
            ' Show sales values for each product as the percentage of its category total.
            ' Total values for each category are displayed as the percentage of the Grand Total value.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.PercentOfParentRow)
'            #End Region ' #PercentOfParentRowTotal
        End Sub

        Private Shared Sub RankLargestToSmallest(ByVal workbook As IWorkbook)
'            #Region "#RankLargestToSmallest"
            Dim worksheet As Worksheet = workbook.Worksheets("Report13")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Amount"), "Rank")
            ' Display the rank of sales values for the "Customer" field, listing the largest item in the field as 1.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.RankDescending, pivotTable.Fields("Customer"))
'            #End Region ' #RankLargestToSmallest
        End Sub

        Private Shared Sub RunningTotalIn(ByVal workbook As IWorkbook)
'            #Region "#RunningTotalIn"
            Dim worksheet As Worksheet = workbook.Worksheets("Report15")
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Access the pivot table by its name in the collection
            Dim pivotTable As PivotTable = worksheet.PivotTables("PivotTable1")

            ' Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Amount"), "Running Total")
            ' Display values for successive items in the "Quarter" field as a running total.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.RunningTotal, pivotTable.Fields("Quarter"))
            ' Specify the number format for the data field.
            dataField.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)"
'            #End Region ' #RunningTotalIn
        End Sub

        Private Shared Sub NumberFormat(ByVal workbook As IWorkbook)
'            #Region "#NumberFormat"
            Dim sourceWorksheet As Worksheet = workbook.Worksheets("Data5")
            Dim worksheet As Worksheet = workbook.Worksheets.Add()
            workbook.Worksheets.ActiveWorksheet = worksheet

            ' Create a pivot table using the cell range "A1:E65" as the data source.
            Dim pivotTable As PivotTable = worksheet.PivotTables.Add(sourceWorksheet("A1:E65"), worksheet("B2"))

            ' Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Category"))
            ' Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields("Product"))

            ' Add the "Amount" field to the data area.
            Dim dataField As PivotDataField = pivotTable.DataFields.Add(pivotTable.Fields("Amount"))
            ' Specify the number format for the data field.
            dataField.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)"
'            #End Region ' #NumberFormat
        End Sub
    End Class
End Namespace
