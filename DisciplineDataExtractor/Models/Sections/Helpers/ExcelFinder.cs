using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DisciplineDataExtractor.Models.Sections.Helpers
{
    public static class ExcelFinder
    {
        /// <summary>
        /// Поиск заданного слова на странице
        /// </summary>
        /// <param name="worksheet">страница для поиска</param>
        /// <param name="target">слово которое ищем</param>
        /// <param name="isRegex">true если хотим передать регекс, иначе false</param>
        /// <returns>адресс ячейки где нашли слово(первый встреченный)</returns>
        /// <exception cref="Exception">нет искомого поля</exception>
        public static string FindCell(IXLWorksheet worksheet, string target, bool isRegex = false)
        {
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ToString();
                        }
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        /// <summary>
        /// ищет один и тот же патерн в документе 2 раза и возвращает их адреса (разные).
        /// </summary>
        public static string[] FindTwoCell(IXLWorksheet worksheet, string target, bool isRegex = false)
        {
            string[] answ = new string[2];
            int count = 0;
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target, RegexOptions.IgnoreCase))
                        {
                            answ[count++] = cell.Address.ToString();
                            if (count == 2) { return answ; }
                        }
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе в количестве 2-ух штук.");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                        {
                            answ[count++] = cell.Address.ToString();
                            if (count == 2) { return answ; }
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        /// <summary>
        /// вложенный поиск ищет слово target2 в ячейках под словом target1
        /// </summary>
        /// <param name="worksheet">страница для поиска</param>
        /// <param name="target1">ищем столбец в котором будем искать</param>
        /// <param name="target2">то что ищем</param>
        /// <param name="isRegex">true если хотим передать регекс, иначе false</param>
        /// <returns>адрес ячейки с target2</returns>
        /// <exception cref="Exception"></exception>
        public static string FindCell(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            // FIX: throw был внутри цикла по строкам — выбрасывался после первой строки.
            // Теперь сначала ищем target1 по всем строкам, потом target2 под ним.
            foreach (var row in worksheet.RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    if (!cell.GetValue<string>().Contains(target1, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var mergedRange = cell.MergedRange() ?? cell.AsRange();
                    var firstColumn = mergedRange.FirstCell().Address.ColumnLetter;
                    var lastColumn = mergedRange.LastCell().Address.ColumnLetter;
                    int startRow = mergedRange.LastCell().Address.RowNumber + 1;
                    int endRow = worksheet.LastRowUsed().RowNumber();
                    var searchRange = worksheet.Range($"{firstColumn}{startRow}:{lastColumn}{endRow}");

                    if (isRegex)
                    {
                        foreach (var cellValue in searchRange.CellsUsed())
                        {
                            string cellForReg = cellValue.GetValue<string>();
                            if (Regex.IsMatch(cellForReg, target2, RegexOptions.IgnoreCase))
                            {
                                return cellValue.Address.ToString();
                            }
                        }
                    }
                    else
                    {
                        foreach (var cellValue in searchRange.CellsUsed())
                        {
                            if (cellValue.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                            {
                                return cellValue.Address.ToString();
                            }
                        }
                    }
                    // target1 найден, но target2 под ним нет — продолжаем искать другой target1
                }
            }
            throw new Exception($"Нет поля {target1}/{target2} в документе {worksheet.Name}");
        }

        /// <summary>
        /// ищем СТОЛБЕЦ в котором содержится искомый патерн
        /// </summary>
        public static string FindColumn(IXLWorksheet worksheet, string target, bool isRegex = false)
        {
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target} в документе");
            }
        }

        /// <summary>
        /// вложенный поиск ищет адрес столбца слова target2 в ячейках под словом target1
        /// </summary>
        public static string FindColumn(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            // FIX: аналогично FindCell — throw был внутри цикла по строкам
            foreach (var row in worksheet.RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    if (!cell.GetValue<string>().Contains(target1, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var mergedRange = cell.MergedRange() ?? cell.AsRange();
                    var firstColumn = mergedRange.FirstCell().Address.ColumnLetter;
                    var lastColumn = mergedRange.LastCell().Address.ColumnLetter;
                    int startRow = mergedRange.LastCell().Address.RowNumber + 1;
                    int endRow = worksheet.LastRowUsed().RowNumber();
                    var searchRange = worksheet.Range($"{firstColumn}{startRow}:{lastColumn}{endRow}");

                    if (isRegex)
                    {
                        foreach (var cellValue in searchRange.CellsUsed())
                        {
                            string cellForReg = cellValue.GetValue<string>();
                            if (Regex.IsMatch(cellForReg, target2, RegexOptions.IgnoreCase))
                            {
                                return cellValue.Address.ColumnLetter.ToString();
                            }
                        }
                    }
                    else
                    {
                        foreach (var cellValue in searchRange.CellsUsed())
                        {
                            if (cellValue.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                            {
                                return cellValue.Address.ColumnLetter.ToString();
                            }
                        }
                    }
                    // target1 найден, но target2 под ним нет — продолжаем искать другой target1
                }
            }
            throw new Exception($"Нет поля {target1}/{target2} в документе {worksheet.Name}");
        }


        public static string FindColumnAnderCell(IXLWorksheet worksheet, IXLCell cell, string target, bool isRegex = false)
        {
            var mergedRange = cell.MergedRange() ?? cell.AsRange();
            var firstColumn = mergedRange.FirstCell().Address.ColumnLetter;
            var lastColumn = mergedRange.LastCell().Address.ColumnLetter;
            int startRow = mergedRange.LastCell().Address.RowNumber + 1;
            int endRow = worksheet.LastRowUsed().RowNumber();
            var searchRange = worksheet.Range($"{firstColumn}{startRow}:{lastColumn}{endRow}");

            if (isRegex)
            {
                foreach (var cellValue in searchRange.CellsUsed())
                {
                    string cellForReg = cellValue.GetValue<string>();
                    if (Regex.IsMatch(cellForReg, target, RegexOptions.IgnoreCase))
                    {
                        return cellValue.Address.ColumnLetter.ToString();
                    }
                }
                throw new Exception($"Нет ПАТЕРНА {target} в документе");
            }
            else
            {
                foreach (var cellValue in searchRange.CellsUsed())
                {
                    if (cellValue.GetValue<string>().Contains(target, StringComparison.OrdinalIgnoreCase))
                    {
                        return cellValue.Address.ColumnLetter.ToString();
                    }
                }
            }
            throw new Exception($"Нет поля {target} в документе {worksheet.Name}");
        }


        public static string FindCellOr(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            if (isRegex)
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target1, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (Regex.IsMatch(cellValue, target2, RegexOptions.IgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target1}, или {target2} в документе {worksheet.Name}");
            }
            else
            {
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target1, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.GetValue<string>().Contains(target2, StringComparison.OrdinalIgnoreCase))
                        {
                            return cell.Address.ColumnLetter.ToString();
                        }
                    }
                }
                throw new Exception($"Нет поля {target1}, или {target2} в документе {worksheet.Name}");
            }
        }

        /// <summary>
        /// FIX: Добавлен отсутствующий метод FindColumnOr — аналог FindCellOr, но возвращает букву столбца.
        /// Ищет target1, если не найден — ищет target2. Возвращает букву столбца первого совпадения.
        /// </summary>
        public static string FindColumnOr(IXLWorksheet worksheet, string target1, string target2, bool isRegex = false)
        {
            // Просто делегируем в FindCellOr — он уже возвращает ColumnLetter
            return FindCellOr(worksheet, target1, target2, isRegex);
        }
    }
}