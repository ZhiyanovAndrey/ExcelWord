using System.Data;
using System.IO;
using System.Text.RegularExpressions;
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

    
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(path)
                        join t in OpenExcelFile.GetTask(path) on p.PersonNumber equals t.PersonNumber
                        select new { Name = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };

            datagrid1.ItemsSource = query.GroupBy(p => p.Name).Select(g => new { Name = g.Key, Count = g.Count() });
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(path)
                        join t in OpenExcelFile.GetTask(path) on p.PersonNumber equals t.PersonNumber 
                        into pt from subb in pt.DefaultIfEmpty()
                        join d in OpenExcelFile.GetDepartment(path) on p.Department equals d.DepartmentId
                        into ptd from subc in ptd.DefaultIfEmpty()
                        group new { p, pt, ptd} by new {p.SurName , p.Department } into g
                        select new { Department = g.Key.Department, g.Key.SurName, Count = g.Count() };
            //into pd from pd
            //join t in OpenExcelFile.GetTask(path) on p.PersonNumber equals t.PersonNumber
            //group new { p, d, t } by t.PersonNumber into g
            //select new { Department = g.Select(x => x.d.Name), Name = g.Select(x => x.p.SurName), Count = g.Count() };

            datagrid1.ItemsSource = query;
        }
        //        from subb in Group1.DefaultIfEmpty()
        //                      join c in Table3 on a.Id equals c.AId into Group2
        //                      from subc in Group2.DefaultIfEmpty()
        //                      group new { a, subb, subc
        //    }
        //    by new { a.Id, a.Name
        //}
        //into g
        //                       select new
        //                              {
        //                                  g.Key.Id,
        //                                  g.Key.Name,
        //                                  SubGroup1Count = g.Count(x => x.subb != null),
        //                                  SubGroup2Count = g.Count(x => x.subc != null)

        private void Button_Click_3(object sender, RoutedEventArgs e)
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

