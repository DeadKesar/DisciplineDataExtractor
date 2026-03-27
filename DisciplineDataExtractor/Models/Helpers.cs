using ClosedXML.Excel;
using DisciplineDataExtractor.Extensions;
using DisciplineDataExtractor.Models.Sections.Helpers;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static DisciplineDataExtractor.Models.Sections.Helpers.ExcelFinder;


namespace DisciplineDataExtractor.Models
{
    public static class Helpers
    {
        private const string WorksheetName = "План";

        public static IDictionary<string, Discipline> GetDisciplines(IXLWorkbook workbook, HierarchicalCheckableElement section, string EducationLevel)
        {
            int semMax = EducationLevel switch
            {
                "Бакалавриат" => 8,
                "Магистратура" => 4,
                "Аспирантура" => 6,
                _ => 8
            };

            var worksheet = workbook.Worksheet(WorksheetName);
            string[] forFactandByPlan = new string[2];
            bool isSpec = false;
            if (EducationLevel == "Специалитет")
            {
                isSpec = true;
                // FindTwoCell возвращает полные адреса ("G5", "H5"), а row.Cell() ожидает букву столбца.
                // Извлекаем только буквенную часть.
                var fullAddresses = FindTwoCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true);
                forFactandByPlan[0] = new string(fullAddresses[0].TakeWhile(char.IsLetter).ToArray());
                forFactandByPlan[1] = new string(fullAddresses[1].TakeWhile(char.IsLetter).ToArray());
            }

            var disciplines = ParseDisciplinesFromSheet(
                ExcelHelpers.GetRowsWithPlus(worksheet),
                worksheet, section, isSpec, forFactandByPlan, semMax);

            if (!disciplines.Any())
            {
                var svodSheet = workbook.Worksheet(WorksheetName + "Свод");
                disciplines = ParseDisciplinesFromSheet(
                    ExcelHelpers.GetRowsWithPlus(svodSheet),
                    worksheet, section, isSpec, forFactandByPlan, semMax);
            }

            return disciplines;
        }

        /// <summary>
        /// Общий метод парсинга дисциплин из набора строк.
        /// Устраняет дублирование кода между "План" и "ПланСвод".
        /// </summary>
        private static Dictionary<string, Discipline> ParseDisciplinesFromSheet(
            IEnumerable<IXLRow> rows,
            IXLWorksheet worksheet,
            HierarchicalCheckableElement section,
            bool isSpec,
            string[] forFactandByPlan,
            int semMax)
        {
            // FIX: row.Cell(string) ожидает БУКВУ столбца ("A", "B", "AC"),
            // а FindCell возвращает полный адрес ("A3"). Поэтому используем FindColumn.
            return rows
                .Select(row => new Discipline
                {
                    Ind = row.Cell(FindColumn(worksheet, "индекс")).GetString(),
                    Name = row.Cell(FindColumn(worksheet, "наименование")).GetString(),
                    Department = row.Cell(FindColumn(worksheet, "закрепленная кафедра", "наименование")).GetString(),
                    Exam = row.Cell(FindColumn(worksheet, "[Э|э]?\\s*[K|к]\\s*[З|з]\\s*[А|а]\\s*[М|м]\\s*[Е|е]\\s*[Н|н]", true)).GetInt(),
                    Credit = row.Cell(FindColumn(worksheet, "зачет")).GetInt(),
                    CreditWithRating = row.Cell(FindColumn(worksheet, "зачет с оц")).GetInt(),
                    Kp = row.Cell(FindColumn(worksheet, "^кп$", true)).GetInt(),
                    Kr = row.Cell(FindColumn(worksheet, "^кр$", true)).GetInt(),
                    Fact = row.Cell(isSpec ? forFactandByPlan[0] : FindColumn(worksheet, "факт")).GetInt(),
                    ByPlan = row.Cell(isSpec ? forFactandByPlan[1] : FindColumnOr(worksheet, "[П|п]?\\s*[О|о]\\s*[П|п]\\s*[Л|л]\\s*[А|а]\\s*[Н|н]s*[У|у]", "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true)).GetInt(),
                    ContactHours = row.Cell(FindColumn(worksheet, "Конт. раб.")).GetInt(),

                    Lec = row.Cell(FindColumn(worksheet, "^лек$", true)).GetInt(),
                    Lab = row.Cell(FindColumn(worksheet, "^лаб$", true)).GetInt(),
                    Pr = row.Cell(FindColumn(worksheet, "^пр$", true)).GetInt(),

                    Control = row.Cell(FindColumn(worksheet, "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                    ZeAtAll = row.Cells(FindColumn(worksheet, "Семестр 1"), FindColumn(worksheet, $"Семестр {semMax}")).Sum(val => val.GetInt()),
                    Parent = section
                })
                .Aggregate(new Dictionary<string, Discipline>(), (dict, discipline) =>
                {
                    string originalName = discipline.Ind;
                    string nameToUse = originalName;
                    int counter = 2;

                    while (dict.ContainsKey(nameToUse))
                    {
                        nameToUse = $"{originalName}{counter}";
                        counter++;
                    }

                    dict[nameToUse] = discipline;
                    return dict;
                });
        }
    }
}