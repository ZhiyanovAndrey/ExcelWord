using ExcelWord.Models;
using System.Data;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;


namespace ExcelWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //private string filename = string.Empty;
        private readonly string _path = @"D:\Data.xlsb";
        private readonly string _wordTemplate = @"D:\Template.doc";

        private WordExporter _wordExporter;


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


            var query = from p in OpenExcelFile.GetPerson(_path)
                        join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                        select new { Name = $"{p.SurName} {p.FirstName}", DepartmentName = d.Name };



            datagrid1.ItemsSource = query.GroupBy(p => p.DepartmentName).Select(g => new { Name = g.Key, Count = g.Count() });


        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(_path)
                        join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                        select new { Name = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };

            datagrid1.ItemsSource = query.GroupBy(p => p.Name).Select(g => new { Name = g.Key, Count = g.Count() });
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(_path)
                        join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                        into pt
                        from subb in pt.DefaultIfEmpty()
                        join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                        into ptd
                        from subc in ptd.DefaultIfEmpty()
                        group new { p, subb, subc } by new { p.SurName, p.Department } into g
                        select new { g.Key.Department, g.Key.SurName, Group = g.Count(x => x.subb != null), Group_ptd = g.Count(x => x.subc != null) };
            //select new { g.Key.Department, g.Key.SurName, Group_pt = g.Count(x => x.subb != null), Group_ptd = g.Count(x => x.subc != null) };

            datagrid1.ItemsSource = query;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var query = from p in OpenExcelFile.GetPerson(_path)
                        join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                        join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                        select new { Отдел = d.Name, ФИО = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };
            //group new { p, d, t } by { p.SurName };
            // Многоуровневая группировка в LINQ?

            datagrid1.ItemsSource = query.GroupBy(q => q.ФИО).Select(g => new { Name = g.Key, Count = g.Count() });
            // GroupBy(x=> new { x.Column1, x.Column2 }, (ключ, группа) => new { Key1 = ключ.Column1, Key2 = ключ.Column2 ,
        }



        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            dynamic excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            WordExporter.WordExport(datagrid1);




            //try
            //{

            //var wordApp = new Word.Application();
            //wordApp.Visible = false;
            //var document = wordApp.Documents.Open(_wordTemplate);
            //document.Activate();
            //Word.Table table = document.Tables[1]; // таблица загруженная из документа

            //table.Cell(1,1).Range.Text = _wordTemplate;

            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show(ex.Message);
            //}





        }
    }





}

