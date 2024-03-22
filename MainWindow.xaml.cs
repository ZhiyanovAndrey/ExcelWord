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
            var persons = GetPerson().ToList();
            var departments = GetDepartment().ToList();


            var query = from p in persons
                        join d in departments on p.Department equals d.DepartmentId
                        select new { Name = $"{p.SurName} {p.FirstName}", DepartmentName = d.Name };



            datagrid1.ItemsSource = query.GroupBy(p => p.DepartmentName).Select(g => new { Name = g.Key, Count = g.Count() }); 
                
               // .Where(g => g.Count() > 1)
               //.Select(g => g.Key); 
        }

            // Загрузить файл Excel
            Workbook wb = new Workbook(@"D:\Data.xlsx");


        public IEnumerable<Person> GetPerson()
        {

            // Получить рабочий лист 1
            using (Worksheet worksheet = wb.Worksheets[1])
            {
                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;

                // Цикл по строкам
                for (int i = 1; i <= rows; i++)
                {
                    var person = new Person
                    {
                        PersonNumber = worksheet.Cells[i, 0].StringValue,
                        SurName = worksheet.Cells[i, 1].StringValue,
                        FirstName = worksheet.Cells[i, 2].StringValue,
                        MiddleName = worksheet.Cells[i, 3].StringValue,
                        Birthday = worksheet.Cells[i, 4].DateTimeValue,
                        Department = worksheet.Cells[i, 5].IntValue,

                    };
                    // И возвращаем его
                    yield return person;
                }
            };
        }


        public IEnumerable<Department> GetDepartment()
        {

            // Получить рабочий лист 2
            using (Worksheet worksheet = wb.Worksheets[2])
            {
                int rows = worksheet.Cells.MaxDataRow;
                for (int i = 1; i <= rows; i++)
                {
                    var department = new Department
                    {
                        DepartmentId = worksheet.Cells[i, 0].IntValue,
                        Name = worksheet.Cells[i, 1].StringValue,
                    };

                    yield return department;
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

