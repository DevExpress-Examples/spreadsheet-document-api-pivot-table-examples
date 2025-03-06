using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraTab;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace SpreadsheetDocServerPivotAPI
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        Workbook workbook = new Workbook();
        public Form1()
        {
            InitializeComponent();
            InitTreeListControl();
            workbook.Options.CalculationMode = WorkbookCalculationMode.Automatic;
        }

        void InitTreeListControl()
        {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }

        void InitData(GroupsOfSpreadsheetExamples examples)
        {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Pivot Calculated Fields"));
            examples.Add(new SpreadsheetNode("Pivot Calculated Item"));
            examples.Add(new SpreadsheetNode("Pivot Fields"));
            examples.Add(new SpreadsheetNode("Pivot Field Groups"));
            examples.Add(new SpreadsheetNode("Pivot Tables"));
            examples.Add(new SpreadsheetNode("Pivot Table Filter"));
            examples.Add(new SpreadsheetNode("Pivot Table Formatting"));
            examples.Add(new SpreadsheetNode("Pivot Table Layout"));
            examples.Add(new SpreadsheetNode("Value Field Settings"));

            #endregion

            #region ExampleNodes
            // Add nodes to the "Calculated Field" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Add Calculated Field", PivotCalculatedFieldActions.AddCalculatedFieldAction));
            examples[0].Groups.Add(new SpreadsheetExample("Modify Calculated Field", PivotCalculatedFieldActions.ModifyCalculatedFieldAction));
            examples[0].Groups.Add(new SpreadsheetExample("Remove Calculated Field", PivotCalculatedFieldActions.RemoveCalculatedFieldAction));

            // Add nodes to the "Calculated Item" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Add Calculated Item", PivotCalculatedItemActions.AddCalculatedItemAction));
            examples[1].Groups.Add(new SpreadsheetExample("Modify Calculated Item", PivotCalculatedItemActions.ModifyCalculatedItemAction));
            examples[1].Groups.Add(new SpreadsheetExample("Remove Calculated Item", PivotCalculatedItemActions.RemoveCalculatedItemAction));

            // Add nodes to the "Pivot Field" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Add Field to Axis", PivotFieldActions.AddFieldToAxisAction));
            examples[2].Groups.Add(new SpreadsheetExample("Insert Field to Axis", PivotFieldActions.InsertFieldToAxisAction));
            examples[2].Groups.Add(new SpreadsheetExample("Move Field Down", PivotFieldActions.MoveFieldDownAction));
            examples[2].Groups.Add(new SpreadsheetExample("Move Field to Axis", PivotFieldActions.MoveFieldToAxisAction));
            examples[2].Groups.Add(new SpreadsheetExample("Move Field Up", PivotFieldActions.MoveFieldUpAction));
            examples[2].Groups.Add(new SpreadsheetExample("Multiple Subtotals", PivotFieldActions.MultipleSubtotalsAction));
            examples[2].Groups.Add(new SpreadsheetExample("Remove Field from Axis", PivotFieldActions.RemoveFieldFromAxisAction));
            examples[2].Groups.Add(new SpreadsheetExample("Sort Field Items", PivotFieldActions.SortFieldItemsAction));
            examples[2].Groups.Add(new SpreadsheetExample("Sort Field Items by Data Field", PivotFieldActions.SortFieldItemsByDataFieldAction));


            // Add nodes to the "Chart Legend" group of examples.
            examples[3].Groups.Add(new SpreadsheetExample("Group Field by Dates", PivotFieldGroupingActions.GroupFieldByDatesAction));
            examples[3].Groups.Add(new SpreadsheetExample("Group Field Items", PivotFieldGroupingActions.GroupFieldItemsAction));
            examples[3].Groups.Add(new SpreadsheetExample("Ungroup Field Items", PivotFieldGroupingActions.UngroupFieldItemsAction));
            examples[3].Groups.Add(new SpreadsheetExample("Ungroup Specific Item", PivotFieldGroupingActions.UngroupSpecificItemAction));

            // Add nodes to the "Pivot Table" group of examples.        
            examples[4].Groups.Add(new SpreadsheetExample("Create Pivot Table from Cache", PivotTableActions.CreatePivotTableFromCacheAction)); 
            examples[4].Groups.Add(new SpreadsheetExample("Create Pivot Table from Range", PivotTableActions.CreatePivotTableFromRangeAction)); 
            examples[4].Groups.Add(new SpreadsheetExample("Change Behavior Options", PivotTableActions.ChangeBehaviorOptionsAction)); 
            examples[4].Groups.Add(new SpreadsheetExample("Change Pivot Table Data Source", PivotTableActions.ChangePivotTableDataSourceAction)); 
            examples[4].Groups.Add(new SpreadsheetExample("Change Pivot Table Location", PivotTableActions.ChangePivotTableLocationAction));
            examples[4].Groups.Add(new SpreadsheetExample("Clear Pivot Table", PivotTableActions.ClearPivotTableAction));
            examples[4].Groups.Add(new SpreadsheetExample("Move Pivot Table to Worksheet", PivotTableActions.MovePivotTableToWorksheetAction));
            examples[4].Groups.Add(new SpreadsheetExample("Remove Pivot Table", PivotTableActions.RemovePivotTableAction));

            // Add nodes to the "Pivot Table Filter" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Set Item Filter", PivotTableFilterActions.SetItemFilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Label Filter", PivotTableFilterActions.SetLabelFilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Date Filter", PivotTableFilterActions.SetDateFilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Multiple Filter", PivotTableFilterActions.SetMultipleFilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Top 10 Filter", PivotTableFilterActions.SetTop10FilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Value Filter", PivotTableFilterActions.SetValueFilterAction));
            examples[5].Groups.Add(new SpreadsheetExample("Set Item Visibility Filter", PivotTableFilterActions.SetItemVisibilityFilterAction));

            // Add nodes to the "Pivot Table Formatting" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Banded Columns", PivotTableFormattingActions.BandedColumnsAction));
            examples[6].Groups.Add(new SpreadsheetExample("Banded Rows", PivotTableFormattingActions.BandedRowsAction));
            examples[6].Groups.Add(new SpreadsheetExample("Change Pivot Table Style", PivotTableFormattingActions.ChangeStylePivotTableAction));
            examples[6].Groups.Add(new SpreadsheetExample("Show Column Headers", PivotTableFormattingActions.ShowColumnHeadersAction));
            examples[6].Groups.Add(new SpreadsheetExample("Show Row Headers", PivotTableFormattingActions.ShowRowHeadersAction));

            // Add nodes to the "Pivot Table Layout" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Column Grand Totals", PivotTableLayoutActions.ColumnGrandTotalsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Data On Rows", PivotTableLayoutActions.DataOnRowsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Hide All Subtotals", PivotTableLayoutActions.HideAllSubtotalsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Insert Blank Rows", PivotTableLayoutActions.InsertBlankRowsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Merge Titles", PivotTableLayoutActions.MergeTitlesAction));
            examples[7].Groups.Add(new SpreadsheetExample("Remove Blank Rows", PivotTableLayoutActions.RemoveBlankRowsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Repeat All Item Labels", PivotTableLayoutActions.RepeatAllItemLabelsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Row Grand Totals", PivotTableLayoutActions.RowGrandTotalsAction));
            examples[7].Groups.Add(new SpreadsheetExample("Set Compact Report Layout", PivotTableLayoutActions.SetCompactReportLayoutAction));
            examples[7].Groups.Add(new SpreadsheetExample("Set Outline Report Layout", PivotTableLayoutActions.SetOutlineReportLayoutAction));
            examples[7].Groups.Add(new SpreadsheetExample("Set Tabular Report Layout", PivotTableLayoutActions.SetTabularReportLayoutAction));
            examples[7].Groups.Add(new SpreadsheetExample("Show All Subtotals", PivotTableLayoutActions.ShowAllSubtotalsAction));

            // Add nodes to the "Value Field Settings" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Change Summary Function", ValueFieldSettingsActions.ChangeSummaryFunctionAction));
            examples[8].Groups.Add(new SpreadsheetExample("Difference From", ValueFieldSettingsActions.DifferenceFromAction));
            examples[8].Groups.Add(new SpreadsheetExample("Number Format", ValueFieldSettingsActions.NumberFormatAction));
            examples[8].Groups.Add(new SpreadsheetExample("Percent Of", ValueFieldSettingsActions.PercentOfAction));
            examples[8].Groups.Add(new SpreadsheetExample("Percent Of Parent Row", ValueFieldSettingsActions.PercentOfParentRowTotalAction));
            examples[8].Groups.Add(new SpreadsheetExample("Rank Largest to Smallest", ValueFieldSettingsActions.RankLargestToSmallestAction));
            examples[8].Groups.Add(new SpreadsheetExample("Running Total In", ValueFieldSettingsActions.RunningTotalInAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }


        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<Workbook> action = example.Action;
            action(workbook);
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile()
        {
            #region #LoadDocumentFromFile
            // Load a workbook from the file.
            workbook.LoadDocument("PivotTableTemplate.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }


        private void SaveDocumentToFile()
        {
            #region #SaveDocumentToFile
            // Save the modified document to the file.
            workbook.SaveDocument("SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
            Process.Start(new ProcessStartInfo("SavedDocument.xlsx") { UseShellExecute = true });
        }

    }
}
