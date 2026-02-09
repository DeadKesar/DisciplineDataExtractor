using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DisciplineWorkProgram.Models.Sections.Helpers;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using static DisciplineWorkProgram.Models.Sections.Helpers.ExcelFinder;


namespace DisciplineWorkProgram.Models
{
    public static class Helpers
    {
        private const string WorksheetName = "План";

        public static IDictionary<string, Discipline> GetDisciplines(IXLWorkbook workbook, HierarchicalCheckableElement section, string EducationLevel)
        {
            int semMax = 0;
            switch (EducationLevel)
            {
                case "Бакалавриат":
                    {
                        semMax = 8;
                        break;
                    }
                case "Магистратура":
                    {
                        semMax = 4;
                        break;
                    }
                case "Аспирантура":
                    {
                        semMax = 6;
                        break;
                    }
                default:
                    {
                        semMax = 8;
                        break;
                    }
            }

            var worksheet = workbook.Worksheet(WorksheetName);
            string[] forFactandByPlan = new string[2];
            bool isSpec = false;
            if (EducationLevel == "Специалитет")
            {
                isSpec = true;
                forFactandByPlan = FindTwoCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true);

            }

            var disciplines = ExcelHelpers.GetRowsWithPlus(worksheet)
                .Select(row => new Discipline
                {
                    Ind = row.Cell(FindCell(worksheet, "индекс")).GetString(),
                    Name = row.Cell(FindCell(worksheet, "наименование")).GetString(), //C
                    Department = row.Cell(FindCell(worksheet, "закрепленная кафедра", "наименование")).GetString(),
                    Exam = row.Cell(FindCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[З|з]\\s*[А|а]\\s*[М|м]\\s*[Е|е]\\s*[Н|н]", true)).GetInt(),
                    Credit = row.Cell(FindCell(worksheet, "зачет")).GetInt(),
                    CreditWithRating = row.Cell(FindCell(worksheet, "зачет с оц")).GetInt(),
                    Kp = row.Cell(FindCell(worksheet, "^кп$", true)).GetInt(),
                    Kr = row.Cell(FindCell(worksheet, "^кр$", true)).GetInt(),
                    Fact = row.Cell(isSpec ? forFactandByPlan[0] : FindCell(worksheet, "факт")).GetInt(),
                    ByPlan = row.Cell(isSpec ? forFactandByPlan[1] : FindCellOr(worksheet, "[П|п]?\\s*[О|о]\\s*[П|п]\\s*[Л|л]\\s*[А|а]\\s*[Н|н]s*[У|у]", "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true)).GetInt(), //экспертное
                    ContactHours = row.Cell(FindCell(worksheet, "Конт. раб.")).GetInt(),
                    Lec = row.Cell(FindCell(worksheet, "Лаб")).GetInt(),
                    Lab = row.Cell(FindCell(worksheet, "^пр$", true)).GetInt(),
                    Pr = row.Cell(FindCell(worksheet, "^ср$", true)).GetInt(),

                    Control = row.Cell(FindCell(worksheet, "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                    ZeAtAll = row.Cells(FindCell(worksheet, "Семестр 1"), FindCell(worksheet, $"Семестр {semMax}")).Sum(val => val.GetInt()),

                    Parent = section
                })
                    .Aggregate(new Dictionary<string, Discipline>(), (dict, discipline) =>
                    {
                        string originalName = discipline.Ind;
                        string nameToUse = originalName;
                        int counter = 2;

                        // Пока ключ уже существует, добавляем суффикс
                        while (dict.ContainsKey(nameToUse))
                        {
                            nameToUse = $"{originalName}{counter}";
                            counter++;
                        }

                        // Добавляем дисциплину с уникальным именем
                        dict[nameToUse] = discipline;

                        return dict;
                    });

            if (!disciplines.Any())
            {
                disciplines = ExcelHelpers.GetRowsWithPlus(workbook.Worksheet(WorksheetName + "Свод"))
                    .Select(row => new Discipline
                    {
                        Name = row.Cell(FindCell(worksheet, "наименование")).GetString(), //C
                        Department = row.Cell(FindCell(worksheet, "закрепленная кафедра", "наименование")).GetString(),
                        Exam = row.Cell(FindCell(worksheet, "[Э|э]?\\s*[K|к]\\s*[З|з]\\s*[А|а]\\s*[М|м]\\s*[Е|е]\\s*[Н|н]", true)).GetInt(),
                        Credit = row.Cell(FindCell(worksheet, "зачет")).GetInt(),
                        CreditWithRating = row.Cell(FindCell(worksheet, "зачет с оц")).GetInt(),
                        Kp = row.Cell(FindCell(worksheet, "^кп$", true)).GetInt(),
                        Kr = row.Cell(FindCell(worksheet, "^кр$", true)).GetInt(),
                        Fact = row.Cell(FindCell(worksheet, "факт")).GetInt(),
                        ByPlan = row.Cell(FindCellOr(worksheet, "[П|п]?\\s*[О|о]\\s*[П|п]\\s*[Л|л]\\s*[А|а]\\s*[Н|н]s*[У|у]", "[Э|э]?\\s*[K|к]\\s*[С|с]\\s*[П|п]\\s*[Е|е]\\s*[Р|р]\\s*[Т|т]\\s*[Н|н]\\s*[О|о]\\s*[Е|е]", true)).GetInt(), //экспертное
                        ContactHours = row.Cell(FindCell(worksheet, "Конт. раб.")).GetInt(),
                        Lec = row.Cell(FindCell(worksheet, "Лаб")).GetInt(),
                        Lab = row.Cell(FindCell(worksheet, "^пр$", true)).GetInt(),
                        Pr = row.Cell(FindCell(worksheet, "^ср$", true)).GetInt(),
                        Ind = row.Cell(FindCell(worksheet, "индекс")).GetText(),
                        Control = row.Cell(FindCell(worksheet, "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                        ZeAtAll = row.Cells(FindCell(worksheet, "Семестр 1"), FindCell(worksheet, "Семестр 8")).Sum(val => val.GetInt()),

                        Parent = section
                    })
                        .Aggregate(new Dictionary<string, Discipline>(), (dict, discipline) =>
                        {
                            string originalName = discipline.Ind;
                            string nameToUse = originalName;
                            int counter = 2;

                            // Пока ключ уже существует, добавляем суффикс
                            while (dict.ContainsKey(nameToUse))
                            {
                                nameToUse = $"{originalName}{counter}";
                                counter++;
                            }

                            // Добавляем дисциплину с уникальным именем
                            dict[nameToUse] = discipline;

                            return dict;
                        });
            }
            //Возможно надо отфильтровать дисциплины. и убрать заголовки


            return disciplines;
        }

      
    }
}
