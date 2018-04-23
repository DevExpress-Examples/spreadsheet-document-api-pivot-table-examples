using DevExpress.Spreadsheet;

namespace SpreadsheetDocServerPivotAPI
{
    class PivotCalculatedFieldActions
    {
        static void AddCalculatedField(IWorkbook workbook)
        {
            #region #AddCalculatedField
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Create a calculated field based on data in the "Sales" field.
            PivotField field = pivotTable.CalculatedFields.Add("=Sales*10%", "Sales Tax");
            // Add the calculated field to the data area and specify the custom field name.
            PivotDataField dataField = pivotTable.DataFields.Add(field, "Total Tax");
            // Specify the number format for the data field.
            dataField.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)";
            #endregion #AddCalculatedField
        }

        static void RemoveCalculatedField(IWorkbook workbook)
        {
            #region #RemoveCalculatedField
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Create a calculated field based on data in the "Sales" field.
            pivotTable.CalculatedFields.Add("=Sales*10%", "Sales Tax");
            // Access the calculated field by its name in the collection.
            PivotField field = pivotTable.CalculatedFields["Sales Tax"];
            // Add the calculated field to the data area. 
            PivotDataField dataField = pivotTable.DataFields.Add(field);
            //Remove the calculated field.
            pivotTable.CalculatedFields.RemoveAt(0);
            #endregion #RemoveCalculatedField
        }

        static void ModifyCalculatedField(IWorkbook workbook)
        {
            #region #ModifyCalculatedField
            Worksheet worksheet = workbook.Worksheets["Report1"];
            workbook.Worksheets.ActiveWorksheet = worksheet;

            // Access the pivot table by its name in the collection.
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            // Create a calculated field based on data in the "Sales" field.
            pivotTable.CalculatedFields.Add("=Sales*10%", "Sales Tax Rate 10");
            // Access the calculated field by its name in the collection.
            PivotField field = pivotTable.CalculatedFields["Sales Tax Rate 10"];
            //Change the formula for the calculated field.
            field.Formula = "=Sales*15%";
            //Change the calculated field name.
            field.Name = "Sales Tax Rate 15";
            //Add the calculated field to the data area and specify the custom field name.
            PivotDataField dataField = pivotTable.DataFields.Add(field, "Total Tax");
            // Specify the number format for the data field.
            dataField.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "" - ""??_);_(@_)";
            #endregion #ModifyCalculatedField
        }
    }
}
