using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    public static class PivotTableActions
    {

        static void CreatePivotTableFromRange(IWorkbook workbook)
        {
            #region #CreateFromRange
            Worksheet sourceWorksheet = workbook.Worksheets["Data1"];
            Worksheet worksheet = workbook.Worksheets.Add();
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a pivot table using the cell range "A1:D41" as the data source.
            PivotTable pivotTable = worksheet.PivotTables.Add(sourceWorksheet["A1:D41"], worksheet["B2"]);

            // Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Category"]);
            // Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Product"]);
            // Add the "Sales" field to the data area.
            pivotTable.DataFields.Add(pivotTable.Fields["Sales"]);
            #endregion #CreateFromRange
        }

        static void CreatePivotTableFromCache(IWorkbook workbook)
        {
            #region #CreateFromPivotCache
            Worksheet worksheet = workbook.Worksheets.Add();
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Create a pivot table based on the specified PivotTable cache.
            PivotCache cache = workbook.Worksheets["Report1"].PivotTables["PivotTable1"].Cache;
            PivotTable pivotTable = worksheet.PivotTables.Add(cache, worksheet["B2"]);

            // Add the "Category" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Category"]);
            // Add the "Product" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["Product"]);
            // Add the "Sales" field to the data area.
            pivotTable.DataFields.Add(pivotTable.Fields["Sales"]);

            // Set the default style for the pivot table.
            pivotTable.Style = workbook.TableStyles.DefaultPivotStyle;

            #endregion #CreateFromPivotCache
        }

        static void RemovePivotTable(IWorkbook workbook)
        {
            #region #RemoveTable
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Remove the pivot table from the collection.
            worksheet.PivotTables.Remove(pivotTable);

            #endregion #RemoveTable
        }
        static void ChangePivotTableLocation(IWorkbook workbook)
        {
            #region #ChangeLocation
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Change the pivot table location.
            worksheet.PivotTables["PivotTable1"].MoveTo(worksheet["A7"]);
            // Refresh the pivot table.
            worksheet.PivotTables["PivotTable1"].Cache.Refresh();

            #endregion #ChangeLocation
        }
        static void MovePivotTableToWorksheet(IWorkbook workbook)
        {
            #region #MoveToWorksheet
            Worksheet worksheet = workbook.Worksheets["Report1"];

            // Create a new worksheet.
            Worksheet targetWorksheet = workbook.Worksheets.Add();

            // Access the pivot table by its name in the collection
            // and move it to the new worksheet.
            worksheet.PivotTables["PivotTable1"].MoveTo(targetWorksheet["B2"]);
            // Refresh the pivot table.
            targetWorksheet.PivotTables["PivotTable1"].Cache.Refresh();

            workbook.Worksheets.ActiveWorksheet = targetWorksheet;
            #endregion #MoveToWorksheet
        }

        static void ChangePivotTableDataSource(IWorkbook workbook)
        {
            #region #ChangeDataSource
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            Worksheet sourceWorksheet = workbook.Worksheets["Data2"];
            // Change the data source of the pivot table.
            pivotTable.ChangeDataSource(sourceWorksheet["A1:H6367"]);

            // Add the "State" field to the row axis area.
            pivotTable.RowFields.Add(pivotTable.Fields["State"]);
            // Add the "Yearly Earnings" field to the data area.
            PivotDataField dataField = pivotTable.DataFields.Add(pivotTable.Fields["Yearly Earnings"]);
            // Calculate the average of the "Yearly Earnings" values for each state.
            dataField.SummarizeValuesBy = PivotDataConsolidationFunction.Average;

            #endregion #ChangeDataSource
        }

        static void ClearPivotTable(IWorkbook workbook)
        {
            #region #ClearTable
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Clear the pivot table.
            worksheet.PivotTables["PivotTable1"].Clear();
            #endregion #ClearTable
        }

        static void ChangeBehaviorOptions(IWorkbook workbook)
        {
            #region #ChangeBehaviorOptions
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;
            worksheet.Columns["B"].WidthInCharacters = 40;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Restrict specific operations for the pivot table.
            PivotBehaviorOptions behaviorOptions = pivotTable.Behavior;
            behaviorOptions.AutoFitColumns = false;
            behaviorOptions.EnableFieldList = false;

            // Refresh the pivot table.
            pivotTable.Cache.Refresh();
            #endregion #ChangeBehaviorOptions
        }
    }
}
