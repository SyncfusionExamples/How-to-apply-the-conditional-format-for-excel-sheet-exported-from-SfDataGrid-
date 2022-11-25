# How to apply the conditional format for excel sheet exported from WPF DataGrid (SfDataGrid)
## About the sample

This example illustrates how to apply the conditional formatting to the exported excel in [WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid).

[WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid) does not provide direct support to the conditional formatting while exporting from the grid to excel sheet. But you can achieve this by applying the conditional formatting to the exported worksheet. 

```c#
 var options = new ExcelExportingOptions();
 options.ExcelVersion = ExcelVersion.Excel2013;
 var excelEngine = sfDataGrid.ExportToExcel(sfDataGrid.View, options);
 var workBook = excelEngine.Excel.Workbooks[0];
 //Apply conditional format to worksheet
 IWorksheet worksheet = workBook.Worksheets[0];
 IConditionalFormats formats = worksheet["F2:F11"].ConditionalFormats;
 IConditionalFormat format = formats.AddCondition();
 format.FormatType = ExcelCFType.DataBar;
 IDataBar dataBar = format.DataBar;
 dataBar.BarColor = Color.Blue;
 workBook.SaveAs("Sample.xlsx");

```

KB article - [How-to-apply-the-conditional-format-for-excel-sheet-exported-from-SfDataGrid-](https://www.syncfusion.com/kb/10994/how-to-apply-the-conditional-format-for-excel-sheet-exported-from-wpf-datagrid-sfdatagrid)

## Requirements to run the demo
Visual Studio 2015 and above versions
