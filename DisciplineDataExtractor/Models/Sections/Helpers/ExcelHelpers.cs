using ClosedXML.Excel;
using NPOI.SS.Formula.Functions;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models.Sections.Helpers
{
    public static class ExcelHelpers
    {
        static string pattern = @"^\+|^[а-яА-Я]\d\.|^[а-яА-Я]{3}\.";
        public static IEnumerable<IXLRow> GetRowsWithPlus(IXLWorksheet worksheet) =>
            worksheet.RowsUsed().Where(row =>
            {
                var cellValue = row.Cell("A").GetString();
                return Regex.IsMatch(cellValue, pattern);
            });

        //row.Cell("A").GetString().Equals("+"));

        //public static IEnumerable<IXLRow> GetRowsWithPractices(IXLWorksheet worksheet) =>
        // 	worksheet.RowsUsed().Where(row => row.Cell("D").GetString().ToLower().Contains("практика"));
    }
}
