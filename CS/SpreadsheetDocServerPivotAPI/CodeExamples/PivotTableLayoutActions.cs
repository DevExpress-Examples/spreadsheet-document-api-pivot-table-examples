using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    public static class PivotTableLayoutActions
    {
        static void ColumnGrandTotals(IWorkbook workbook)
        {
            #region #ColumnGrandTotals
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Add the "Region" field to the column axis area. 
            pivotTable.ColumnFields.Add(pivotTable.Fields["Region"]);

            // Hide grand totals for columns.
            pivotTable.Layout.ShowColumnGrandTotals = false;
            #endregion #ColumnGrandTotals
        }

        static void RowGrandTotals(IWorkbook workbook)
        {
            #region #RowGrandTotals
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Add the "Region" field to the column axis area. 
            pivotTable.ColumnFields.Add(pivotTable.Fields["Region"]);

            // Hide grand totals for rows.
            pivotTable.Layout.ShowRowGrandTotals = false;
            #endregion #RowGrandTotals
        }

        static void DataOnRows(IWorkbook workbook)
        {
            #region #MultipleDataFields
            Worksheet worksheet = workbook.Worksheets["Report2"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display value fields in separate columns.
            pivotTable.Layout.DataOnRows = false;
            #endregion #MultipleDataFields
        }
        static void MergeTitles(IWorkbook workbook)
        {
            #region #MergeTitles
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display the pivot table in the tabular form.
            pivotTable.Layout.SetReportLayout(PivotReportLayout.Tabular);
            // Merge and center cells with labels. 
            pivotTable.Layout.MergeTitles = true;
            #endregion #MergeTitles
        }

        static void ShowAllSubtotals(IWorkbook workbook)
        {
            #region #ShowAllSubtotals
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Show all subtotals at the top of each group.
            pivotTable.Layout.ShowAllSubtotals(true);

            #endregion #ShowAllSubtotals
        }

        static void HideAllSubtotals(IWorkbook workbook)
        {
            #region #HideAllSubtotals
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Hide subtotals at the top of each group.
            pivotTable.Layout.HideAllSubtotals();
            #endregion #HideAllSubtotals
        }

        static void SetCompactReportLayout(IWorkbook workbook)
        {
            #region #CompactReportLayout
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display the pivot table in the compact form.
            pivotTable.Layout.SetReportLayout(PivotReportLayout.Compact);
            #endregion #CompactReportLayout
        }

        static void SetOutlineReportLayout(IWorkbook workbook)
        {
            #region #OutlineReportLayout
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display the pivot table in the outline form.
            pivotTable.Layout.SetReportLayout(PivotReportLayout.Outline);
            #endregion #OutlineReportLayout
        }

        static void SetTabularReportLayout(IWorkbook workbook)
        {
            #region #TabularReportLayout
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display the pivot table in the tabular form.
            pivotTable.Layout.SetReportLayout(PivotReportLayout.Tabular);
            #endregion #TabularReportLayout
        }

        static void RepeatAllItemLabels(IWorkbook workbook)
        {
            #region #RepeatAllItemLabels
            Worksheet worksheet = workbook.Worksheets["Report5"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Display repeated column labels.
            pivotTable.Layout.RepeatAllItemLabels(true);
            #endregion #RepeatAllItemLabels
        }

        static void InsertBlankRows(IWorkbook workbook)
        {
            #region #InsertBlankRows
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Insert a blank row after each group of items.
            pivotTable.Layout.InsertBlankRows();

            #endregion #InsertBlankRows
        }

        static void RemoveBlankRows(IWorkbook workbook)
        {
            #region #RemoveBlankRows
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Insert a blank row after each group of items.
            pivotTable.Layout.InsertBlankRows();

            // Remove a blank row after each group of items.
            pivotTable.Layout.RemoveBlankRows();
            #endregion #RemoveBlankRows
        }
    }
}
