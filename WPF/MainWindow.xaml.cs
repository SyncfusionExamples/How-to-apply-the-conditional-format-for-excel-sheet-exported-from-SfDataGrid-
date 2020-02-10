using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.UI.Xaml.Grid;
using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.UI.Xaml.Grid.Helpers;
using Syncfusion.XlsIO;

namespace SfDataGrid_Sample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
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
        }
    }
}
