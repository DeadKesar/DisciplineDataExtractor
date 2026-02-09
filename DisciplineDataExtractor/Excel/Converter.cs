using System;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;   // Для старых .xls (HSSF)
using NPOI.SS.UserModel;    // Общий интерфейс для HSSF/XSSF
using NPOI.XSSF.UserModel;  // Для .xlsx (XSSF)

namespace DisciplineDataExtractor.Excel
{
    public class Converter
    {
        private const string ConvertedFilesPath = "Converted/";
        private const string Extension = ".xlsx";

        public static Stream Convert2(string path)//, out string newPath) Void
        {
            //newPath = GetConvertedFilePath(path);
            if (!Directory.Exists(ConvertedFilesPath)) Directory.CreateDirectory(ConvertedFilesPath);
            return ConvertToXlsx(path);
            //ConvertXlsToXlsx(path).Write(File.Create(newPath));
        }

        //создаём в памяти почищенный учебный план
        public static Stream Convert(string path)
        {
            var newPlan = new NpoiMemoryStream() { AllowClose = false };
            var newPath = GetConvertedFilePath(path);

            ConvertXlsToXlsx2(path).Write(newPlan);
            newPlan.AllowClose = true;

            return newPlan;
        }

        private static XSSFWorkbook ConvertXlsToXlsx(string path)
        {
            // Путь к исходному файлу
            string xlsFilePath = path;
            // Путь к файлу, в который будет конвертирован XLS
            string xlsxFilePath = GetConvertedFilePath(path);


            // Открываем старый Excel файл (.xls)
            using (var fileStream = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read))
            {
                // Загружаем старый файл (.xls) с помощью HSSF (для .xls файлов)
                HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileStream);

                // Создаем новый Excel файл (.xlsx)
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

                // Переносим все листы из старого файла в новый
                for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
                {
                    ISheet sheet = hssfWorkbook.GetSheetAt(i);
                    ISheet newSheet = xssfWorkbook.CreateSheet(sheet.SheetName);

                    // Копируем строки и ячейки
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            IRow newRow = newSheet.CreateRow(rowIndex);
                            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                            {
                                ICell cell = row.GetCell(cellIndex);
                                ICell newCell = newRow.CreateCell(cellIndex);

                                if (cell != null)
                                {
                                    newCell.SetCellValue(cell.ToString());
                                }
                            }
                        }
                    }
                }
                // Сохраняем новый файл (.xlsx)
                using (var fs = new FileStream(xlsxFilePath, FileMode.Create, FileAccess.Write))
                {
                    xssfWorkbook.Write(fs);
                }
                return xssfWorkbook;
            }

        }

        private static XSSFWorkbook ConvertXlsToXlsx2(string path)
        {
            using var inputStream = File.OpenRead(path);
            // var workbookIn = new HSSFWorkbook(inputStream);
            var workbookIn = new XSSFWorkbook(inputStream);
            var workbookOut = new XSSFWorkbook();

            for (var i = 0; i < workbookIn.NumberOfSheets; i++)
            {
                var sheetIn = workbookIn.GetSheetAt(i);
                var sheetOut = workbookOut.CreateSheet(sheetIn.SheetName);
                var rowEnumerator = sheetIn.GetEnumerator();
                while (rowEnumerator.MoveNext())
                {
                    // var rowIn = (HSSFRow)rowEnumerator.Current;
                    var rowIn = (XSSFRow)rowEnumerator.Current;
                    // if (rowIn == null || rowIn.IsHidden) continue;
                    if (rowIn == null) continue;
                    var rowOut = sheetOut.CreateRow(rowIn.RowNum);
                    CopyRowProperties(rowOut, rowIn);
                }
            }
            return workbookOut;
        }

        private static void CopyRowProperties(IRow rowOut, IRow rowIn)
        {
            rowOut.RowNum = rowIn.RowNum;

            using var cellEnumerator = rowIn.GetEnumerator();

            while (cellEnumerator.MoveNext())
            {
                var cellIn = cellEnumerator.Current;
                if (cellIn == null) continue;
                var cellOut = rowOut.CreateCell(cellIn.ColumnIndex, cellIn.CellType);
                CopyCellProperties(cellOut, cellIn);
            }
        }

        private static void CopyCellProperties(ICell cellOut, ICell cellIn)
        {
            switch (cellIn.CellType)
            {
                case CellType.Blank:
                    break;
                case CellType.Boolean:
                    cellOut.SetCellValue(cellIn.BooleanCellValue);
                    break;
                case CellType.Error:
                    cellOut.SetCellValue(cellIn.ErrorCellValue);
                    break;
                case CellType.Formula:
                    cellOut.SetCellFormula(cellIn.CellFormula);
                    break;
                case CellType.Numeric:
                    cellOut.SetCellValue(cellIn.NumericCellValue);
                    break;
                case CellType.String:
                    cellOut.SetCellValue(cellIn.StringCellValue);
                    break;
                case CellType.Unknown:
                    break;
                default:
                    return;
            }
        }




        private static string GetConvertedFilePath(string path) =>
            Path.GetFullPath("dwp/" + Path.GetFileNameWithoutExtension(path) + Extension);



        /// <summary>
        /// Конвертирует входной .xls или .xlsx в новый .xlsx и возвращает Stream.
        /// Использует "двухшаговый" подход, чтобы обойти проблему закрытия потока в DotNetCore.NPOI.
        /// </summary>
        public static Stream ConvertToXlsx(string inputPath)
        {
            // 1) Открываем входной файл (xls или xlsx)
            IWorkbook workbookIn = OpenWorkbook(inputPath);
            if (workbookIn == null)
                throw new FileNotFoundException("Не удалось открыть файл или файл не является Excel.", inputPath);

            // 2) Создаём выходной XSSFWorkbook
            var workbookOut = new XSSFWorkbook();

            // 3) Готовим словарь для стилей
            var styleMap = new Dictionary<short, ICellStyle>();

            // 4) Копируем все листы
            for (int i = 0; i < workbookIn.NumberOfSheets; i++)
            {
                ISheet sheetIn = workbookIn.GetSheetAt(i);
                if (sheetIn == null) continue;

                ISheet sheetOut = workbookOut.CreateSheet(sheetIn.SheetName);
                CopySheet(sheetIn, sheetOut, workbookIn, workbookOut, styleMap);
            }

            // 5) Записываем workbookOut в промежуточный MemoryStream (который DotNetCore.NPOI может закрыть)
            byte[] resultBytes;
            using (var msTemp = new MemoryStream())
            {
                workbookOut.Write(msTemp);
                // После Write(...) DotNetCore.NPOI может закрыть msTemp,
                // поэтому сразу делаем ToArray.
                resultBytes = msTemp.ToArray();
            }

            // 6) Создаём НОВЫЙ MemoryStream из массива байт
            var msFinal = new MemoryStream(resultBytes, writable: false);
            // Перематываем на начало
            msFinal.Position = 0;

            // Возвращаем «живой» поток
            return msFinal;
        }

        /// <summary>
        /// Определяет формат файла по расширению и открывает как HSSFWorkbook или XSSFWorkbook.
        /// </summary>
        private static IWorkbook OpenWorkbook(string path)
        {
            if (!File.Exists(path))
                return null;

            var ext = Path.GetExtension(path).ToLowerInvariant();
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);

            if (ext == ".xls")
                return new HSSFWorkbook(fs);    // старый HSSF
            else
                return new XSSFWorkbook(fs);    // новый XSSF
        }

        /// <summary>
        /// Копирует один лист (Sheet) целиком:
        /// - ширины и скрытость столбцов,
        /// - объединённые ячейки,
        /// - строки (высота, скрытость),
        /// - ячейки (значения, формулы, стили, шрифты).
        /// </summary>
        private static void CopySheet(
            ISheet sheetIn,
            ISheet sheetOut,
            IWorkbook workbookIn,
            XSSFWorkbook workbookOut,
            Dictionary<short, ICellStyle> styleMap)
        {
            // 1) Определяем максимальное количество столбцов (чтобы скопировать ширины)
            int maxColumn = 0;
            for (int rowIndex = sheetIn.FirstRowNum; rowIndex <= sheetIn.LastRowNum; rowIndex++)
            {
                IRow rowIn = sheetIn.GetRow(rowIndex);
                if (rowIn != null && rowIn.LastCellNum > maxColumn)
                    maxColumn = rowIn.LastCellNum;
            }

            // 2) Копируем ширины и скрытость столбцов
            for (int col = 0; col < maxColumn; col++)
            {
                sheetOut.SetColumnWidth(col, sheetIn.GetColumnWidth(col));
                sheetOut.SetColumnHidden(col, sheetIn.IsColumnHidden(col));
            }

            // 3) Копируем объединённые ячейки
            int mergedCount = sheetIn.NumMergedRegions;
            for (int i = 0; i < mergedCount; i++)
            {
                var mergedRegion = sheetIn.GetMergedRegion(i);
                sheetOut.AddMergedRegion(mergedRegion);
            }

            // 4) Копируем строки
            for (int rowIndex = sheetIn.FirstRowNum; rowIndex <= sheetIn.LastRowNum; rowIndex++)
            {
                IRow rowIn = sheetIn.GetRow(rowIndex);
                if (rowIn == null) continue;

                IRow rowOut = sheetOut.CreateRow(rowIndex);

                // Высота строки
                rowOut.Height = rowIn.Height;
                // Скрытость строки
                rowOut.ZeroHeight = rowIn.ZeroHeight;

                // 5) Копируем ячейки
                for (int cellIndex = rowIn.FirstCellNum; cellIndex < rowIn.LastCellNum; cellIndex++)
                {
                    ICell cellIn = rowIn.GetCell(cellIndex);
                    if (cellIn == null)
                        continue;

                    ICell cellOut = rowOut.CreateCell(cellIndex);
                    CopyCell(cellIn, cellOut, workbookIn, workbookOut, styleMap);
                }
            }
        }

        /// <summary>
        /// Копирует одну ячейку (стиль + значение).
        /// </summary>
        private static void CopyCell(
            ICell cellIn,
            ICell cellOut,
            IWorkbook workbookIn,
            XSSFWorkbook workbookOut,
            Dictionary<short, ICellStyle> styleMap)
        {
            // 1) Копируем стиль
            var styleIn = cellIn.CellStyle;
            if (styleIn != null)
            {
                short styleIndex = styleIn.Index;

                if (!styleMap.TryGetValue(styleIndex, out ICellStyle cachedStyle))
                {
                    cachedStyle = CopyStyleManually(workbookOut, styleIn, workbookIn);
                    styleMap[styleIndex] = cachedStyle;
                }

                cellOut.CellStyle = cachedStyle;
            }

            // 2) Копируем тип и значение
            switch (cellIn.CellType)
            {
                case CellType.Blank:
                    cellOut.SetCellType(CellType.Blank);
                    break;

                case CellType.Boolean:
                    cellOut.SetCellValue(cellIn.BooleanCellValue);
                    break;

                case CellType.Error:
                    cellOut.SetCellErrorValue(cellIn.ErrorCellValue);
                    break;

                case CellType.Formula:
                    // Сохраняем «текст» формулы (без пересчёта)
                    cellOut.SetCellFormula(cellIn.CellFormula);
                    break;

                case CellType.Numeric:
                    cellOut.SetCellValue(cellIn.NumericCellValue);
                    break;

                case CellType.String:
                    cellOut.SetCellValue(cellIn.StringCellValue);
                    break;

                default:
                    cellOut.SetCellValue(cellIn.ToString());
                    break;
            }
        }

        /// <summary>
        /// Полностью вручную копирует свойства стиля из styleIn (HSSF или XSSF) в новый XSSFCellStyle workbookOut.
        /// </summary>
        private static ICellStyle CopyStyleManually(
            XSSFWorkbook workbookOut,
            ICellStyle styleIn,
            IWorkbook workbookIn)
        {
            if (styleIn == null)
                return null;

            // Создаём новый стиль
            ICellStyle styleOut = workbookOut.CreateCellStyle();

            // Копируем базовые поля стиля
            styleOut.Alignment = styleIn.Alignment;
            styleOut.VerticalAlignment = styleIn.VerticalAlignment;
            styleOut.BorderLeft = styleIn.BorderLeft;
            styleOut.BorderRight = styleIn.BorderRight;
            styleOut.BorderTop = styleIn.BorderTop;
            styleOut.BorderBottom = styleIn.BorderBottom;

            styleOut.LeftBorderColor = styleIn.LeftBorderColor;
            styleOut.RightBorderColor = styleIn.RightBorderColor;
            styleOut.TopBorderColor = styleIn.TopBorderColor;
            styleOut.BottomBorderColor = styleIn.BottomBorderColor;

            styleOut.FillPattern = styleIn.FillPattern;
            styleOut.FillForegroundColor = styleIn.FillForegroundColor;
            styleOut.FillBackgroundColor = styleIn.FillBackgroundColor;

            styleOut.DataFormat = styleIn.DataFormat;
            styleOut.WrapText = styleIn.WrapText;
            styleOut.Indention = styleIn.Indention;
            styleOut.Rotation = styleIn.Rotation;
            styleOut.IsLocked = styleIn.IsLocked;
            styleOut.IsHidden = styleIn.IsHidden;

            // Копируем шрифт
            short fontIndex = styleIn.FontIndex;
            IFont fontIn = workbookIn.GetFontAt(fontIndex);
            if (fontIn != null)
            {
                IFont fontOut = FindOrCreateFont(workbookOut, fontIn);
                styleOut.SetFont(fontOut);
            }

            return styleOut;
        }

        /// <summary>
        /// Ищет или создаёт в workbookOut (XSSFWorkbook) шрифт, похожий на fontIn.
        /// </summary>
        private static IFont FindOrCreateFont(XSSFWorkbook workbookOut, IFont fontIn)
        {
            // Пытаемся найти уже существующий такой же шрифт
            for (short i = 0; i < workbookOut.NumberOfFonts; i++)
            {
                IFont existing = workbookOut.GetFontAt(i);
                if (IsSameFont(existing, fontIn))
                    return existing;
            }

            // Если не нашли, создаём новый
            IFont newFont = workbookOut.CreateFont();

            newFont.Boldweight = fontIn.Boldweight;
            newFont.Color = fontIn.Color;
            newFont.FontHeight = fontIn.FontHeight;
            newFont.FontName = fontIn.FontName;
            newFont.IsItalic = fontIn.IsItalic;
            newFont.Underline = fontIn.Underline;
            newFont.TypeOffset = fontIn.TypeOffset;
            newFont.Charset = fontIn.Charset;

            return newFont;
        }

        /// <summary>
        /// Проверяем, одинаковые ли основные свойства двух шрифтов.
        /// </summary>
        private static bool IsSameFont(IFont f1, IFont f2)
        {
            if (f1.Boldweight != f2.Boldweight) return false;
            if (f1.Color != f2.Color) return false;
            if (f1.FontHeight != f2.FontHeight) return false;
            if (f1.FontName != f2.FontName) return false;
            if (f1.IsItalic != f2.IsItalic) return false;
            if (f1.Underline != f2.Underline) return false;
            if (f1.TypeOffset != f2.TypeOffset) return false;
            if (f1.Charset != f2.Charset) return false;
            return true;
        }
    }
}

