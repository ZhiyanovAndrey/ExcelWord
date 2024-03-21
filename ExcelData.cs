using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace ExcelWord
{
    public class ExcelData
    {
        // Табельный номер	Фамилия	Имя 	Отчество	Дата рождения	Отдел

        public string Id { get; set; }
        public string SurName { get; set; }
        public string FirstName { get; set; }
        public string? MiddleName { get; set; }
       
        public DateTime Birthday { get; set; }
        public int Department { get; set; }

        public ExcelData() { }

       
        }
    }
