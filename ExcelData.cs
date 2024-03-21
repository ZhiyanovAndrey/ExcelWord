using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWord
{
    public class ExcelData
    {
        // Табельный номер	Фамилия	Имя 	Отчество	Дата рождения	Отдел

        public int Id { get; set; }
        public string SurName { get; set; }
        public string FirstName { get; set; }
        public string? MiddleName { get; set; }
        public DateTime Birthday { get; set; }
        public int Department { get; set; }

        public ExcelData() { }

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
