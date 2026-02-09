using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DisciplineWorkProgram.Extensions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static DisciplineDataExtractor.Word.Helpers.Tables;
using static DisciplineWorkProgram.Models.Sections.Helpers.Competencies;
using System;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Spreadsheet;
using DisciplineWorkProgram.Models.Sections.Helpers;

namespace DisciplineWorkProgram.Models.Sections
{
    public class Section : HierarchicalCheckableElement //Section - направление
    {
        public string Name => SectionDictionary.ContainsKey("WaySection") ? SectionDictionary["WaySection"] : "";
        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Disciplines.Values;

        private readonly string _compListPath;
        private readonly string _competenciesMatrixPath;

        //Содержит значения Section. Не свойства, поскольку закладки находятся как словарь и проще
        //использовать Section как словарь
        public IDictionary<string, string> SectionDictionary { get; set; }
        public IDictionary<string, Discipline> Disciplines { get; private set; }
        public IDictionary<string, Competence> Competencies { get; set; }
        //Ключ - название дисциплины, значение - список кодов компетенций
        public IDictionary<string, List<string>> DisciplineCompetencies { get; set; }
        public static IDictionary<string, string> CompetenceClassifiers = new Dictionary<string, string>
        {
            ["УК"] = "Универсальные компетенции (УК)",
            ["ОПК"] = "Общепрофессиональные компетенции (ОПК)",
            ["ПК"] = "Профессиональные компетенции (ПК)"
        };
        private Stream? _compListStream;   // поле для хранения потока
        private Stream? _compMatrixStream; 

        //перегрузка конструктора
        public Section(Stream compListStream, Stream compMatrixStream)
        {
            // Сохраняем потоки для последующего использования в LoadCompetenciesData
            _compListStream = compListStream ?? throw new ArgumentNullException(nameof(compListStream));
            _compMatrixStream = compMatrixStream ?? throw new ArgumentNullException(nameof(compMatrixStream));

            SectionDictionary = new Dictionary<string, string>();
            Competencies = new Dictionary<string, Competence>();
            DisciplineCompetencies = new Dictionary<string, List<string>>();
        }
        public Section(string competenciesListPath, string competenciesMatrixPath)
            : this(File.OpenRead(competenciesListPath), File.OpenRead(competenciesMatrixPath)) { }

        public void LoadDataFromPlan(Stream plan)
        {
            plan.Seek(0, SeekOrigin.Begin);
            using var workbook = new XLWorkbook(plan);
            LoadSection(workbook);
        }

        public void LoadDataFromPlan(string path)
        {
            using var stream = File.OpenRead(path);
            LoadDataFromPlan(stream);
        }

        public void LoadCompetenciesData()
        {
            // Используем сохранённые потоки из конструктора
            if (_compListStream == null || _compMatrixStream == null)
                throw new InvalidOperationException("Потоки компетенций не инициализированы. Используйте конструктор с Stream.");

            // Перематываем на начало (на случай многократного вызова)
            _compListStream.Seek(0, SeekOrigin.Begin);
            _compMatrixStream.Seek(0, SeekOrigin.Begin);

            // Загрузка списка компетенций
            using var compListDoc = WordprocessingDocument.Open(_compListStream, false);
            LoadCompetencies(compListDoc);

            // Загрузка матрицы
            using var compMatrixDoc = WordprocessingDocument.Open(_compMatrixStream, false);
            LoadCompetenciesMatrix(compMatrixDoc);
        }

        //Короче, обяз. часть и другие. Их по-идее надо отделять. Может, в коммент как доп. поле поместить
        //к дисциплине или типа того. Но это надо. Наверное.
        //Допустим, здесь все равно
        private void LoadCompetenciesMatrix(WordprocessingDocument document)
        {
            foreach (var table in GetTables(document))
            {
                var headers = GetHeaders(table).ToArray();    //Получить заголовки таблиц
                                                              //По строкам ориентир. Одну пропускаем, так как это заголовки
                foreach (var row in table.Descendants<TableRow>().Skip(1).ToArray())
                {
                    if (row.Descendants<TableCell>().Count() < 2) continue; //Если повторно некоторый заголовок

                    var cells = row.Descendants<TableCell>().ToArray();
                    //var disc = cells[0].Elements<Paragraph>().Single().InnerText; //название дисциплины в первой ячейке
                    var disc = cells[0].InnerText.TrimStart();

                    if (!DisciplineCompetencies.ContainsKey(disc))
                        DisciplineCompetencies[disc] = new List<string>();
                    //Если заголовок не код компетенции или ячейка пуста, то пропускаем
                    for (var i = 1; i < headers.Length; i++)
                    {
                        if (!RegexPatterns.Competence.IsMatch(headers[i]) ||
                            i - (headers.Length - cells.Length) < 0 ||
                            string.IsNullOrWhiteSpace(cells[i-(headers.Length - cells.Length)].InnerText))//string.IsNullOrWhiteSpace(cells[i].Elements<Paragraph>().Single().InnerText))
                            continue;

                        DisciplineCompetencies[disc].Add(headers[i]);
                    }
                }
            }
        }

        private void LoadSection(IXLWorkbook workbook)
        {
            //var regex = new Regex("(?<=\").*(?=\")");
            var worksheet = workbook.Worksheet("Титул");
            SectionDictionary["EducationLevel"] = worksheet.Cell(ExcelFinder.FindCell(worksheet, "квалификация", false)).Value.ToString().ToLower().Replace("квалификация:", "").Trim();
            switch (SectionDictionary["EducationLevel"])
            {
                case "бакалавр":
                    {
                        SectionDictionary["EducationLevel"] = "Бакалавриат";
                        break;
                    }
                case "магистр":
                    {
                        SectionDictionary["EducationLevel"] = "Магистратура";
                        break;
                    }
                case "аспирант":
                    {
                        SectionDictionary["EducationLevel"] = "Аспирантура";
                        break;
                    }
                case "специалист":
                    {
                        SectionDictionary["EducationLevel"] = "Специалитет";
                        break;
                    }

                default:
                    break;
            }
            SectionDictionary["WayCode"] = worksheet.Cell(ExcelFinder.FindCell(worksheet, "\\d\\d.\\d\\d.\\d\\d$", true)).Value.ToString();
            SectionDictionary["EducationForm"] = worksheet.Cell(ExcelFinder.FindCell(worksheet, "форма обучения")).Value.ToString().Replace("Форма обучения: ", "");

            if (SectionDictionary["EducationLevel"] == "Специалитет")
            {
                SectionDictionary["WayName"] = worksheet.Cell(ExcelFinder.FindCell(worksheet, "Специальность:", false)).Value.ToString().Replace("Специальность:", "").Trim();
                SectionDictionary["WaySection"] = worksheet.Cell(ExcelFinder.FindTwoCell(worksheet, "Специализация", false)[1]).Value.ToString();
            }
            else
            {
                //B18 - сложная строка, требуется разложение
                var matches = RegexPatterns.WayNameSection.Matches(worksheet.Cell(ExcelFinder. FindCell(worksheet, "направление подготовки")).Value.ToString());
                SectionDictionary["WayName"] = matches[0].Value;
                SectionDictionary["WaySection"] = matches[1].Value; //Профиль

            }
            Disciplines = DisciplineWorkProgram.Models.Helpers.GetDisciplines(workbook, this, SectionDictionary["EducationLevel"]);
            LoadDetailedDisciplineData(workbook);
        }



        private void LoadDetailedDisciplineData(IXLWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets.Where(sheet => sheet.Name.StartsWith("Курс")))
            {
                int i = 0;
                int cur = 0;
                int count = 1;
                foreach (var row in worksheet.RowsUsed().Where(row => int.TryParse(row.Cell(ExcelFinder.FindColumn(worksheet, "№")).GetString(), out _))
                    .Concat(worksheet.RowsUsed().Where(row =>
                        row.Cell(ExcelFinder.FindColumn(worksheet, "наименование")).GetString().ToLower().ContainsAny("практика", "аттестация", "квалифик")))
                    )
                {

                    var discipline = row.Cell(ExcelFinder.FindColumn(worksheet, "Индекс", true)).GetString();
                    if (string.IsNullOrWhiteSpace(discipline))
                        discipline = row.Cell("E").GetString(); //вроде не актуально

                    if (!Disciplines.ContainsKey(discipline)) continue;
                    //Изменить на трайпарс после дебага

                    HashSet<int> set = new HashSet<int>();
                    foreach (int ind in Disciplines[discipline].Exam.ToString().Select(ch => int.Parse(ch.ToString())).ToArray())
                        set.Add(ind);
                    foreach (int ind in Disciplines[discipline].Credit.ToString().Select(ch => int.Parse(ch.ToString())).ToArray())
                        set.Add(ind);
                    foreach (int ind in Disciplines[discipline].CreditWithRating.ToString().Select(ch => int.Parse(ch.ToString())).ToArray())
                        set.Add(ind);
                    foreach (int ind in Disciplines[discipline].Kr.ToString().Select(ch => int.Parse(ch.ToString())).ToArray())
                        set.Add(ind);
                    foreach (int ind in Disciplines[discipline].Kp.ToString().Select(ch => int.Parse(ch.ToString())).ToArray())
                        set.Add(ind);

                    if (discipline.Contains("Б3"))
                    {
                        switch (SectionDictionary["EducationLevel"])
                        {
                            case "Бакалавриат":
                                {
                                    set.Add(8);
                                    break;
                                }
                            case "Магистратура":
                                {
                                    set.Add(4);
                                    break;
                                }
                            case "Аспирантура":
                                {
                                    set.Add(6);
                                    break;
                                }
                            case "Специалитет":
                                {
                                    set.Add(11);
                                    break;
                                }

                            default:
                                break;
                        }
                    }
                    string[] semestrs = ExcelFinder.FindTwoCell(worksheet, "семестр");
                    var semester = 0;
                    bool isGood = int.TryParse(RegexPatterns.DigitInString.Match(worksheet.Cell(semestrs[0]).GetString()).Value, out semester);
                    if (!isGood) semester = ((cur + 1)) * 2;
                    string[] academChas = ExcelFinder.FindTwoCell(worksheet, "Академических");


                    int.TryParse(row.Cell(ExcelFinder.FindColumn(worksheet, "№")).GetString(), out cur);
                    if (cur < i)
                    {
                        count++;
                    }
                    i = cur;
                    if (count < (semester + 1) / 2)
                    {
                        continue;
                    }

                    if (set.Contains(semester))
                    {
                        var details = new DisciplineDetails
                        {
                            Semester = semester.ToString(),
                            Monitoring = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetString(),
                            Contact = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[А|а]\\s*[К|к]\\s*[Т|т]", true)).GetInt(),
                            Lec = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^лек$", true)).GetInt(),
                            Lab = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^лаб$", true)).GetInt(),
                            Pr = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^пр$", true)).GetInt(),
                            Ind = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^ср$", true)).GetInt(),
                            Control = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[0]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                            Ze = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[0]), @"^з\.е\.$", true)).GetInt()
                        };

                        if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
                            Disciplines[discipline].Details.Add(semester, details);
                    }
                    
                    isGood = int.TryParse(RegexPatterns.DigitInString.Match(worksheet.Cell(semestrs[1]).GetString()).Value, out semester);
                    if (!isGood) semester = ((cur + 1)) * 2 + 1;
                    if (set.Contains(semester))
                    {
                        var details = new DisciplineDetails
                        {
                            Semester = semester.ToString(),
                            Monitoring = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetString(),
                            Contact = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[А|а]\\s*[К|к]\\s*[Т|т]", true)).GetInt(),
                            Lec = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^лек$", true)).GetInt(),
                            Lab = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^лаб$", true)).GetInt(),
                            Pr = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^пр$", true)).GetInt(),
                            Ind = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^ср$", true)).GetInt(),
                            Control = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(academChas[1]), "^[К|к]?\\s*[О|о]\\s*[Н|н]\\s*[Т|т]\\s*[Р|р]\\s*[О|о]\\s*[Л|л]", true)).GetInt(),
                            Ze = row.Cell(ExcelFinder.FindColumnAnderCell(worksheet, worksheet.Cell(semestrs[1]), @"^з\.е\.$", true)).GetInt()
                        };

                        if (!Disciplines[discipline].Details.ContainsKey(semester) && !details.IsHollow)
                            Disciplines[discipline].Details.Add(semester, details);
                    }
                }
            }
        }

        private void LoadCompetencies(WordprocessingDocument document)
        {
            var competencies = ParseCompetencies(document).ToArray();
            var regex = RegexPatterns.CompetenceName2;
            //var regex = new Regex(@"^(УК-[\dЗ]+(\.\d+)*|ОПК-[\dЗ]+(\.\d+)*|ПК-[\dЗ]+(\.[\dЗ]+)*)\b");
            //Составление набора ключей-компетенций 
            foreach (var competency in competencies)
            {
                var match = regex.Match(competency);
                if (match.Success)
                {
                    var key = match.Value.Replace(" ", "").Replace("З", "3");


                    if (!Competencies.ContainsKey(key))
                    {
                        // Если ключа ещё нет, создаём новую компетенцию
                        Competencies[key] = new Competence { Name = competency };
                    }
                    else
                    {
                        // Если ключ уже есть, добавляем строку в список компетенций
                        Competencies[key].Competencies.Add(competency);
                    }
                }
            }


        }

        public IEnumerable<string> GetCheckedDisciplinesNames =>
            Disciplines
                .Where(d => d.Value.IsChecked)
                .Select(kv => kv.Key);

        public IEnumerable<string> GetAnyDisciplinesNames =>
            Disciplines
                .Where(d => true)
                .Select(kv => kv.Key);

    }
}