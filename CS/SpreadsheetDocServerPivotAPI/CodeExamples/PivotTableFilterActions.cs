using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    public static class PivotTableFilterActions
    {
        static void SetItemFilter(IWorkbook workbook)
        {
            #region #ItemFilter
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Show the first item in the "Product" field.
            pivotTable.Fields[1].ShowSingleItem(0);

            //Show all items in the "Product" field (the default option).
            //pivotTable.Fields[1].ShowAllItems();
            #endregion #ItemFilter
        }

        static void SetItemVisibilityFilter(IWorkbook workbook)
        {
            #region #ItemVisibility
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Access items of the "Product" field.
            PivotItemCollection pivotFieldItems = pivotTable.Fields[1].Items;

            // Hide the first item in the "Product" field.
            pivotFieldItems[0].Visible = false;
            #endregion #ItemVisibility
        }

        static void SetLabelFilter(IWorkbook workbook)
        {
            #region #LabelFilter
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Access the "Region" field.
            PivotField field = pivotTable.Fields[0];
            // Filter the "Region" field by text to display sales data for the "South" region.  
            pivotTable.Filters.Add(field, PivotFilterType.CaptionEqual, "South");
            #endregion #LabelFilter
        }

        static void SetValueFilter(IWorkbook workbook)
        {
            #region #ValueFilter
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Access the "Product" field.
            PivotField field = pivotTable.Fields[1];
            // Filter the "Product" field to display products with grand total sales between $6000 and $13000.
            pivotTable.Filters.Add(field, pivotTable.DataFields[0], PivotFilterType.ValueBetween, 6000, 13000);
            #endregion #ValueFilter
        }

        static void SetTop10Filter(IWorkbook workbook)
        {
            #region #Top10Filter
            Worksheet worksheet = workbook.Worksheets["Report4"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Access the "Product" field.
            PivotField field = pivotTable.Fields[1];
            // Filter the "Product" field to display two products with the lowest sales.
            PivotFilter filter = pivotTable.Filters.Add(field, pivotTable.DataFields[0], PivotFilterType.Count, 2);
            filter.Top10Type = PivotFilterTop10Type.Bottom;
            #endregion #Top10Filter
        }

        static void SetDateFilter(IWorkbook workbook)
        {
            #region #DateFilter
            Worksheet worksheet = workbook.Worksheets["Report6"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Access the "Date" field.
            PivotField field = pivotTable.Fields[0];
            // Filter the "Date" field to display sales for the second quarter.
            pivotTable.Filters.Add(field, PivotFilterType.SecondQuarter);
            #endregion #DateFilter
        }

        static void SetMultipleFilter(IWorkbook workbook)
        {
            #region #MultipleFilters
            Worksheet worksheet = workbook.Worksheets["Report6"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Allow multiple filters for a field.
            pivotTable.Behavior.AllowMultipleFieldFilters = true;

            // Filter the "Date" field to display sales for the second quarter.
            PivotField field1 = pivotTable.Fields[0];
            pivotTable.Filters.Add(field1, PivotFilterType.SecondQuarter);

            // Add the second filter to the "Date" field to display two days with the lowest sales.  
            PivotFilter filter = pivotTable.Filters.Add(field1, pivotTable.DataFields[0], PivotFilterType.Count, 2);
            filter.Top10Type = PivotFilterTop10Type.Bottom;
            #endregion #MultipleFilters
        }
    }
}

