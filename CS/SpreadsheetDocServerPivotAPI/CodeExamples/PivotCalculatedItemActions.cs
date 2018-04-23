using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    class PivotCalculatedItemActions
    {
        static void AddCalculatedItem(IWorkbook workbook)
        {
            #region #AddCalculatedItem
            Worksheet worksheet = workbook.Worksheets["Report10"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Access the pivot field by its name in the collection.
            PivotField field = pivotTable.Fields["State"];

            // Add calculated items to the "State" field.
            field.CalculatedItems.Add("=Arizona+California+Colorado", "West Total");
            field.CalculatedItems.Add("=Illinois+Kansas+Wisconsin", "Midwest Total");
            #endregion #AddCalculatedItem
        }

        static void RemoveCalculatedItem(IWorkbook workbook)
        {
            #region #RemoveCalculatedItem
            Worksheet worksheet = workbook.Worksheets["Report7"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Access the pivot field by its name in the collection.
            PivotField field = pivotTable.Fields["Customer"];

            // Add a calculated item to the "Customer" field.
            field.CalculatedItems.Add("='Big Foods'*110%", "Big Foods Sales Plan");

            //Remove the calculated item by its index from the collection.
            field.CalculatedItems.RemoveAt(0);
            #endregion #RemoveCalculatedItem
        }

        static void ModifyCalculatedItem(IWorkbook workbook)
        {
            #region #ModifyCalculatedItem
            Worksheet worksheet = workbook.Worksheets["Report7"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];

            // Access the pivot field by its name in the collection.
            PivotField field = pivotTable.Fields["Customer"];

            // Add a calculated item to the "Customer" field.
            PivotItem item = field.CalculatedItems.Add("='Big Foods'*110%", "Big Foods Sales Plan");

            //Change the formula for the calculated item.
            item.Formula = "='Big Foods'*115%";
            #endregion #ModifyCalculatedItem
        }
    }
}
