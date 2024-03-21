using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Aspose.Cells;
using ExcelDataReader;
using Microsoft.Win32;
namespace ExcelWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //private string filename = string.Empty;



        public MainWindow()
        {
            InitializeComponent();
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
            OpenExcelFile(@"D:\Data.xlsx");
        }

        public static void OpenExcelFile(string path)
        {
            // Загрузить файл Excel
            Workbook wb = new Workbook(path);

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;

            // Перебрать все рабочие листы
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {

                // Получить рабочий лист, используя его индекс
                Worksheet worksheet = collection[worksheetIndex];

                // Печать имени рабочего листа
                Console.WriteLine("Worksheet: " + worksheet.Name);

                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                // Цикл по строкам
                for (int i = 0; i < rows; i++)
                {

                    // Перебрать каждый столбец в выбранной строке
                    for (int j = 0; j < cols; j++)
                    {
                        // Значение ячейки Pring
                        Console.Write(worksheet.Cells[i, j].Value + " | ");
                    }
                    // Распечатать разрыв строки
                    Console.WriteLine(" ");
                }
            }

        }




        //FileStream fileStream=File.Open(path,FileMode.Open,FileAccess.Read);
        //IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream);  // приводим поток к интерфейсу

        //DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration() // создадим бд
        //{
        //    ConfigureDataTable=(x) => new ExcelDataTableConfiguration()
        //    {
        //        UseHeaderRow = true   // считываем верхнюю строку с названием колонок
        //    }
        //});

        //// присвоим
        //datagrid1.ItemsSource=db.Tables;

    }
}
