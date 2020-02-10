using Syncfusion.UI.Xaml.Grid;
using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace SfDataGrid_Sample
{
    public class ViewModel
    {
        private ObservableCollection<OrderInfo> _orders;
        public ObservableCollection<OrderInfo> Orders
        {
            get { return _orders; }
            set { _orders = value; }
        }

        private ICommand exportCommand;

        public ICommand ExportCommand
        {
            get
            {
                return exportCommand;
            }
            set
            {
                exportCommand = value;
            }
        }

        public ViewModel()
        {
            _orders = new ObservableCollection<OrderInfo>();
            this.GenerateOrders();
            ExportCommand = new RelayCommand(OnExportClicked);
        }

        private void GenerateOrders()
        {
            _orders.Add(new OrderInfo(1001, "Maria Anders", "Germany", "ALFKI", "Berlin", 12));
            _orders.Add(new OrderInfo(1002, "Ana Trujilo", "Mexico", "ANATR", "Mexico D.F.", 89));
            _orders.Add(new OrderInfo(1003, "Antonio Moreno", "Mexico", "ANTON", "Mexico D.F.", 78));
            _orders.Add(new OrderInfo(1004, "Thomas Hardy", "UK", "AROUT", "London", 45));
            _orders.Add(new OrderInfo(1005, "Christina Berglund", "Sweden", "BERGS", "Lula", 23));
            _orders.Add(new OrderInfo(1006, "Hanna Moos", "Germany", "BLAUS", "Mannheim", 34));
            _orders.Add(new OrderInfo(1007, "Frederique Citeaux", "France", "BLONP", "Strasbourg", 6));
            _orders.Add(new OrderInfo(1008, "Martin Sommer", "Spain", "BOLID", "Madrid", 56));
            _orders.Add(new OrderInfo(1009, "Laurence Lebihan", "France", "BONAP", "Marseille", 90));
            _orders.Add(new OrderInfo(1010, "Elizabeth Lincoln", "Canada", "BOTTM", "Tsawassen", 78));
        }

        private void OnExportClicked(object obj)
        {
            var grid = obj as SfDataGrid;
            var options = new ExcelExportingOptions();
            options.ExcelVersion = ExcelVersion.Excel2013;
            var excelEngine = grid.ExportToExcel(grid.View, options);
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
