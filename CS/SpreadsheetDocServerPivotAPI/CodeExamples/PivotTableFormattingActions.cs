using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    public static class PivotTableFormattingActions
    {

        static void ChangeStylePivotTable(IWorkbook workbook)
        {
            #region #SetStyle
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Set the pivot table style.
            pivotTable.Style = workbook.TableStyles[BuiltInPivotStyleId.PivotStyleDark7];
            #endregion #SetStyle
        }


        static void BandedColumns(IWorkbook workbook)
        {
            #region #BandedColumns
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Apply the banded column formatting to the pivot table. 
            pivotTable.BandedColumns = true;
            #endregion #BandedColumns
        }

        static void BandedRows(IWorkbook workbook)
        {
            #region #BandedRows
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Apply the banded row formatting to the pivot table. 
            pivotTable.BandedRows = true;
            #endregion #BandedRows
        }

        static void ShowColumnHeaders(IWorkbook workbook)
        {
            #region #ColumnHeaders
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Remove formatting from column headers.
            pivotTable.ShowColumnHeaders = false;
            #endregion #ColumnHeaders
        }

        static void ShowRowHeaders(IWorkbook workbook)
        {
            #region #RowHeaders
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Add the "Region" field to the column axis area. 
            pivotTable.ColumnFields.Add(pivotTable.Fields["Region"]);
            // Remove formatting from row headers.
            pivotTable.ShowRowHeaders = false;
            #endregion #RowHeaders
        }
    }
}
