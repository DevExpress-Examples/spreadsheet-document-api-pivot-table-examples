using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    public static class ValueFieldSettingsActions
    {
        static void ChangeSummaryFunction(IWorkbook workbook)
        {
            #region #ChangeSummaryFunction
            Worksheet sourceWorksheet = workbook.Worksheets["Data5"];
            Worksheet worksheet = workbook.Worksheets.Add();
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a pivot table using the cell range "A1:E65" as the data source.
            PivotTable pivotTable = worksheet.PivotTables.Add(sourceWorksheet["A1:E65"], worksheet["B2"]);

            // Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Category"]);
            // Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Product"]);

            // Add the "Amount" field to the data area.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Amount"]);
            // Use the "Average" function to summarize values in the data field.
            dataField.SummarizeValuesBy = PivotDataConsolidationFunction.Average;
            // Specify the number format for the data field.
            dataField.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)";
            #endregion #ChangeSummaryFunction
        }

        static void DifferenceFrom(IWorkbook workbook)
        {
            #region #DifferenceFrom
            Worksheet worksheet = workbook.Worksheets["Report14"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Access the data field by its index in the collection.
            PivotDataField dataField = pivotTable.DataFields[0];
            // Display the difference in product sales between the current quarter and the previous quarter.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.Difference, pivotTable.Fields["Quarter"], PivotBaseItemType.Previous);
            #endregion #DifferenceFrom
        }

        static void PercentOf(IWorkbook workbook)
        {
            #region #PercentOf
            Worksheet worksheet = workbook.Worksheets["Report14"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Access the data field by its index in the collection.
            PivotDataField dataField = pivotTable.DataFields[0];
            // Select the base field ("Quarter"). 
            PivotField baseField = pivotTable.Fields["Quarter"];
            // Select the base item ("Q1"). 
            PivotItem baseItem = baseField.Items[0];
            // Show values as the percentage of the value of the base item in the base field. 
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.Percent, baseField, baseItem);
            #endregion #PercentOf
        }

        static void PercentOfParentRowTotal(IWorkbook workbook)
        {
            #region #PercentOfParentRowTotal
            Worksheet worksheet = workbook.Worksheets["Report16"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Amount"], "% of Parent Row Total");
            // Show sales values for each product as the percentage of its category total.
            // Total values for each category are displayed as the percentage of the Grand Total value.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.PercentOfParentRow);
            #endregion #PercentOfParentRowTotal
        }

        static void RankLargestToSmallest(IWorkbook workbook)
        {
            #region #RankLargestToSmallest
            Worksheet worksheet = workbook.Worksheets["Report13"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Amount"], "Rank");
            // Display the rank of sales values for the "Customer" field, listing the largest item in the field as 1.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.RankDescending, pivotTable.Fields["Customer"]);
            #endregion #RankLargestToSmallest
        }

        static void RunningTotalIn(IWorkbook workbook)
        {
            #region #RunningTotalIn
            Worksheet worksheet = workbook.Worksheets["Report15"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Add the "Amount" field to the data area for the second time and assign the custom name to the field.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Amount"], "Running Total");
            // Display values for successive items in the "Quarter" field as a running total.
            dataField.ShowValuesWithCalculation(PivotShowValuesAsType.RunningTotal, pivotTable.Fields["Quarter"]);
            // Specify the number format for the data field.
            dataField.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)";
            #endregion #RunningTotalIn
        }

        static void NumberFormat(IWorkbook workbook)
        {
            #region #NumberFormat
            Worksheet sourceWorksheet = workbook.Worksheets["Data5"];
            Worksheet worksheet = workbook.Worksheets.Add();
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a pivot table using the cell range "A1:E65" as the data source.
            PivotTable pivotTable = worksheet.PivotTables.Add(sourceWorksheet["A1:E65"], worksheet["B2"]);

            // Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Category"]);
            // Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Product"]);

            // Add the "Amount" field to the data area.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Amount"]);
            // Specify the number format for the data field.
            dataField.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)";
            #endregion #NumberFormat
        }
    }
}
