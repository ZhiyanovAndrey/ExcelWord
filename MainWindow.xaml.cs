using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Aspose.Cells;
using ExcelDataReader;
using ExcelWord.Models;
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
            var data = OpenExcelFile(@"D:\Data.xlsx").OrderBy(x => x.SurName).ToList();
            datagrid1.ItemsSource = data;
        }

        public IEnumerable<Person> OpenExcelFile(string path)
        {
            // Загрузить файл Excel
            using (Workbook wb = new Workbook(path)) 

            // Получить все рабочие листы
            using (Worksheet worksheet = wb.Worksheets[1])
            {




                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                // Цикл по строкам
                for (int i = 1; i <= rows; i++)
                {
                    var data = new Person
                    {
                        PersonNumber = worksheet.Cells[i, 0].StringValue,
                        SurName = worksheet.Cells[i, 1].StringValue,
                        FirstName = worksheet.Cells[i, 2].StringValue,
                        MiddleName = worksheet.Cells[i, 3].StringValue,
                        Birthday = worksheet.Cells[i, 4].DateTimeValue,
                        Department = worksheet.Cells[i, 5].IntValue,

                    };
                    // И возвращаем его
                    yield return data;
                }
            };
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

