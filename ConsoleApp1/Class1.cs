using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using Ganss.Excel;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;

namespace ConsoleApp1
{
    public class Author
    {
        public int AuthorId { get; set; }
        public string Name { get; set; }
        public string Country { get; set; }

        public static void Example1()
        {
            var authors = new List<Author>
        {
            new Author { AuthorId = 1, Name = "Carson Alexander", Country = "US" },
            new Author { AuthorId = 2, Name = "Meredith Alonso", Country = "UK" },
            new Author { AuthorId = 3, Name = "Arturo Anand", Country = "Canada" },
            new Author { AuthorId = 4, Name = "Gytis Barzdukas", Country = "UK" },
            new Author { AuthorId = 5, Name = "Yan Li", Country = "Japan" },
        };

            var excelMapper = new ExcelMapper();
            excelMapper.Save(@"D:\authors.xlsx", authors);
        }
    }
}
   





