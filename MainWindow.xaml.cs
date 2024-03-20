using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelDataReader;
using Microsoft.Win32;
namespace ExcelWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string filename = string.Empty;
    


        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{
            //    OpenFileDialog ofd = new OpenFileDialog
            //    {
            //        Filter = "";
            //    };  

            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show(ex.Message);
            //}
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenExceleFile(@"D:\Data.xlsx");
        }

        private void OpenExceleFile(string path)
        {
            FileStream fileStream=File.Open(path,FileMode.Open,FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream);  // приводим поток к интерфейсу

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration() // создадим бд
            {
                ConfigureDataTable=(x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true   // считываем верхнюю строку с названием колонок
                }
            });

            // присвоим
            datagrid1.ItemsSource=db.Tables;

        }
    }
}