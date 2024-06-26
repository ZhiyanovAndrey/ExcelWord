﻿using ExcelWord.Models;
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
            try
            {

                var query = from p in OpenExcelFile.GetPerson(_path)
                            join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                            select new { Name = $"{p.SurName} {p.FirstName}", DepartmentName = d.Name };



                datagrid1.ItemsSource = query.GroupBy(p => p.DepartmentName).Select(g => new { Name = g.Key, Count = g.Count() });

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                var query = from p in OpenExcelFile.GetPerson(_path)
                            join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                            select new
                            {
                                Name = $"{p.SurName.Trim()} {p.FirstName.Trim().First()}" +
                            $". {p.MiddleName.FirstOrDefault()}.",
                                TaskName = t.TaskId
                            };

                // количество задач у сотрудников
                datagrid1.ItemsSource = query.GroupBy(p => p.Name).Select(g => new { Name = g.Key, Count = g.Count() });

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
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
            try
            {
                var query = from p in OpenExcelFile.GetPerson(_path)
                            join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                            join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                            select new { Отдел = d.Name, ФИО = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };

                // количество задач у сотрудников по отделам
                var query1 = query.GroupBy(q => new { q.Отдел, q.ФИО }).Select(g => new { g.Key.Отдел, g.Key.ФИО, Count = g.Count() });
                var query2 = query1.GroupBy(q => new { q.Отдел }).Select(g => new { g.Key.Отдел, Count = g.Sum(x => x.Count) });

                datagrid1.ItemsSource = query2;


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }



        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            //dynamic excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            //WordExporter.WordExport(datagrid1);




            try
            {
                Word.Application app = new Word.Application();
                var doc = app.Documents.Add(Visible:true);

                doc.Save();
                doc.Close();
                app.Quit();

                


                //var wordApp = new Word.Application();
                //wordApp.Visible = false;
                //var document = wordApp.Documents.Open(_wordTemplate);
                //document.Activate();
                //Word.Table table = document.Tables[1]; // таблица загруженная из документа

                //table.Cell(1, 1).Range.Text = _wordTemplate;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
               
            }





        }
    }





}

