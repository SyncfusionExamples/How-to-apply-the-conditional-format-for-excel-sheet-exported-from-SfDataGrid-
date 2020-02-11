# How-to-apply-the-conditional-format-for-excel-sheet-exported-from-SfDataGrid-
## About the sample
This example illustrates how to apply the conditional formatting to the exported excel in WPF DataGrid 

DataGrid does not support the conditional formatting while exporting to excel. But this can be achieved by applying the conditional formatting to the exported worksheet. 

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
## Requirements to run the demo
Visual Studio 2015 and above versions
