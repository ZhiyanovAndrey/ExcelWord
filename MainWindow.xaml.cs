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

        const string path = @"D:\Data.xlsb";

        private void Button_Click(object sender, RoutedEventArgs e)
        {


            var query = from p in OpenExcelFile.GetPerson(path)
                        join d in OpenExcelFile.GetDepartment(path) on p.Department equals d.DepartmentId
                        select new { Name = $"{p.SurName} {p.FirstName}", DepartmentName = d.Name };



            datagrid1.ItemsSource = query.GroupBy(p => p.DepartmentName).Select(g => new { Name = g.Key, Count = g.Count() }); 
                
               // .Where(g => g.Count() > 1)
               //.Select(g => g.Key); 
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(path)
                        //join d in OpenExcelFile.GetDepartment(path) on p.Department equals d.DepartmentId
                        join t in OpenExcelFile.GetTask(path) on p.PersonNumber equals t.PersonNumber   
                        select new { Name = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };

            datagrid1.ItemsSource = query;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Workbook wb = new Workbook(path);

            foreach (Worksheet worksheet in wb.Worksheets)
            {
                MessageBox.Show(worksheet.Name);
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

