using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using static DisciplineWorkProgram.Word.Helpers.Ooxml;
using static DisciplineWorkProgram.Word.Helpers.Tables;
using DisciplineWorkProgram.Models.Sections;
using DocumentFormat.OpenXml;
using DisciplineWorkProgram.Models.Sections.Helpers;
using System.Text;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Models
{
    public class Dwp
    {
        public Dwp(Section section)
        {
            Section = section;
        }

        private Section Section { get; }

        //Должно обрабатывать только 1 дисциплину, чтобы "масштабировать" без доп. кода
        public void MakeDwp(string templatePath, string dwpDir, string discipline, Employee employes)
        {


            using var doc = WordprocessingDocument.CreateFromTemplate(templatePath, true);
            var bookmarkMap = GetBookmarks(doc, "Autofill");
            //bool isSafe = false;

            WriteSectionData(bookmarkMap, doc);
            WriteDisciplineData(bookmarkMap, discipline, doc);
            WriteEmploesData(bookmarkMap, discipline, doc, employes);
            WriteRequirements(bookmarkMap, discipline, doc);
            if (dwpDir != "vkr/")
                WriteCompetenciesTable(bookmarkMap, discipline, doc); //заполняет табличку компетенций
            if (dwpDir != "vkr/")
                WriteDisciplinePartitionTable(bookmarkMap, discipline, doc);
            if (dwpDir == "dwp/")
                WritePracticleClassTable(bookmarkMap, discipline, doc); //блок
            if (dwpDir != "vkr/")
                WriteSemesters(bookmarkMap, discipline, doc);
            WriteCompetencies(bookmarkMap, discipline, doc);//записываем компетенции в самом начале
            WriteYear(bookmarkMap, doc);
            // Не реализовано занесение данных по дисциплине
            if (dwpDir != "vkr/")
                WriteLaboriousnessTable(bookmarkMap, discipline, doc);
            if (dwpDir == "vkr/")
                WriteLabVkrTable(bookmarkMap, discipline, doc);
            if (dwpDir == "dwp/")
                WriteLaboratiesClassTable(bookmarkMap, discipline, doc);//блок
            if (dwpDir == "dwp/")
                CheckKurs(bookmarkMap, discipline, doc, templatePath, dwpDir, employes);

            SaveDoc(doc, dwpDir, Section.Disciplines[discipline].Name);
            doc.Dispose();

        }

        private void WriteSectionData(IDictionary<string, BookmarkStart> bookmarkMap, WordprocessingDocument doc)
        {
            foreach (var (key, bookmark) in bookmarkMap)
            {
                var actualKey = key.Substring(0, key.Length - 1);

                try
                {
                    if (Section.SectionDictionary.ContainsKey(actualKey))
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Section.SectionDictionary[actualKey];
                }
                catch
                {
                    Console.WriteLine(key);
                    Environment.Exit(1);
                }
            }
        }

        private void WriteDisciplineData(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline))
                return;

            int countForPartFormForPartners = 0;
            foreach (var (key, bookmark) in bookmarkMap)
            {
                //Надежда на то, что понадобится нумерация в рамках только одноразрядного числа
                var actualKey = key.Substring(0, key.Length - 1);

                var d = Section.Disciplines[discipline].Props.ContainsKey(actualKey);

                if (actualKey == "Discipline")
                {
                    if (Section.Disciplines[discipline].Props.ContainsKey(actualKey))
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Section.Disciplines[discipline].Props["Name"];
                    continue;
                }
                //часть, формируемая участниками образовательных отношений
                //Обязательная часть
                //AutofillPartType1

                if (actualKey == "PartType")
                {
                    if (countForPartFormForPartners == 0)
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Regex.IsMatch(discipline, @".О.") ? "Обязательная часть." : "Часть, формируемая участниками образовательных отношений.";
                    else
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = Regex.IsMatch(discipline, @".О.") ? "обязательной части" : "части, формируемой участниками образовательных отношений";
                    countForPartFormForPartners++;
                    continue;
                }

                if (Section.Disciplines[discipline].Props.ContainsKey(actualKey))
                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = Section.Disciplines[discipline].Props[actualKey];
            }
        }

        private void WriteEmploesData(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc, Employee employes)
        {
            if (!Section.Disciplines.ContainsKey(discipline))
                return;
            foreach (var (key, bookmark) in bookmarkMap)
            {
                var actualKey = key.Substring(0, key.Length - 1);

                switch (actualKey)
                {
                    case "PositionKaf":
                        if (Section.Disciplines[discipline].Props["Department"] == "")
                            continue;
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees[Section.Disciplines[discipline].Props["Department"]]["position"];
                        continue;
                    case "PositionKafForDoc":
                        if (Section.Disciplines[discipline].Props["Department"] == "")
                            continue;
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees[Section.Disciplines[discipline].Props["Department"]]["nameForDoc"];
                        continue;

                    case "PositionKafName":
                        if (Section.Disciplines[discipline].Props["Department"] == "")
                            continue;
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees[Section.Disciplines[discipline].Props["Department"]]["FIO"];
                        continue;

                    case "PositionUmu":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Начальник учебно-методического управления ДСиРОД"]["position"];
                        continue;
                    case "PositionUmuName":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Начальник учебно-методического управления ДСиРОД"]["FIO"];
                        continue;
                    case "PositionBib":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Директор научно-технической библиотеки"]["position"];
                        continue;
                    case "PositionBibName":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Директор научно-технической библиотеки"]["FIO"];
                        continue;
                    case "PositionUitp":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Начальник управления информационно-технической поддержки ДЦТ"]["position"];
                        continue;
                    case "PositionUitpName":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Начальник управления информационно-технической поддержки ДЦТ"]["FIO"];
                        continue;
                    case "PositionInst":
                        if (Section.Disciplines[discipline].Props["Department"] == "" || Section.Disciplines[discipline].Props["Department"] == "Управление по организации проектного обучения")
                            continue;
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                                .First(elem => elem.Text.Contains("Autofill" + actualKey))
                                .Text = employes.Employees[employes.Employees[Section.Disciplines[discipline].Props["Department"]]["institut"]]["position"];
                        continue;
                    case "PositionInstName":
                        if (Section.Disciplines[discipline].Props["Department"] == "" || Section.Disciplines[discipline].Props["Department"] == "Управление по организации проектного обучения")
                            continue;
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                                .First(elem => elem.Text.Contains("Autofill" + actualKey))
                                .Text = employes.Employees[employes.Employees[Section.Disciplines[discipline].Props["Department"]]["institut"]]["FIO"];
                        continue;
                    case "PositionOdimp":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                                .First(elem => elem.Text.Contains("Autofill" + actualKey))
                                .Text = employes.Employees["Проректор по образовательной деятельности и молодежной политике"]["position"];
                        continue;
                    case "PositionOdimpName":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                                .First(elem => elem.Text.Contains("Autofill" + actualKey))
                                .Text = employes.Employees["Проректор по образовательной деятельности и молодежной политике"]["FIO"];
                        continue;
                    case "PredWay":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                                .First(elem => elem.Text.Contains("Autofill" + actualKey))
                                .Text = employes.Employees[Section.SectionDictionary["WayCode"]]["FIO"];
                        continue;
                       

                }
            }
        }

        private void WriteCompetenciesTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;
            //AutofillCompetenciesTable1
            var table = FindElementsByBookmark<Table>(bookmarkMap["CompetenciesTable1"], 2, doc).First();

            foreach (var competence in Section.DisciplineCompetencies[Section.Disciplines[discipline].Name])
            {
                if (!Section.Competencies.ContainsKey(competence)) continue;

                var row = new TableRow();
                row.AppendChild(GetTableCellByString(Section.Competencies[competence].Name));
                var relatedCompetencies = Section.Competencies
                    .Where(kvp => kvp.Key.StartsWith(competence + "."))
                    .ToList();
                var cell = new TableCell();

                var paragraph = new Paragraph();

                foreach (var s in relatedCompetencies)
                {
                    // Добавляем текст в параграф
                    paragraph.AppendChild(new Run(new Text(s.Value.Name)));
                    // Добавляем перенос строки
                    paragraph.AppendChild(new Run(new Break()));
                }

                // Убираем последний Break, если нужно
                if (paragraph.LastChild is Run lastRun && lastRun.LastChild is Break)
                {
                    lastRun.RemoveChild(lastRun.LastChild);
                }

                // Создаём ячейку и добавляем в неё параграф
                var tableCell = new TableCell(paragraph);
                row.AppendChild(tableCell);

                table.AppendChild(row);

            }
        }

        private void WriteDisciplinePartitionTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            var rows = new List<TableRow>();

            foreach (var (semester, _) in Section.Disciplines[discipline].Details)
            {
                var row = new TableRow();
                row.Append(
                    GetTableCellByString(""),
                    GetTableCellByString(""),
                    new TableCell(
                        new Paragraph(
                            new Run(
                                new Text(semester.ToString())))
                        {
                            ParagraphProperties = new ParagraphProperties(
                                new Justification { Val = JustificationValues.Center })
                        }
                    ));
                row.Append(
                    GetTableCellsByStrings("",
                    Section.Disciplines[discipline].Details[semester].Lec.ToString(),
                    Section.Disciplines[discipline].Details[semester].Lab.ToString(),
                    Section.Disciplines[discipline].Details[semester].Pr.ToString(),
                    Section.Disciplines[discipline].Details[semester].Ind.ToString(),
                    "",
                    Section.Disciplines[discipline].Details[semester].Monitoring.ToString()));

                rows.Add(row);
            }

            rows.Add(new TableRow(
                GetTableCellsByStrings(
                    "",
                    "Итого",
                    "",
                    "",
                    Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Lec).ToString(),
                    Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Lab).ToString(),
                    Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Pr).ToString(),
                    Section.Disciplines[discipline].Details.Values.Sum(detail => detail.Ind).ToString(),
                    "",
                    "")));
            rows.Add(new TableRow(
                GetTableCellsByStrings(
                    "",
                    "Промежуточный контроль",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    Section.Disciplines[discipline].Details
                        .FirstOrDefault(elem =>
                            elem.Key == Section.Disciplines[discipline].Details.Keys.Max())
                        .Value?.Monitoring ?? "НЕТ")));

            FindElementsByBookmark<Table>(bookmarkMap["DisciplinePartitionTable1"], 2, doc)
                .First()
                .Append(rows);
        }

        private void WriteSemesters(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name)) return;

            FindElementsByBookmark<Text>(bookmarkMap["Semester1"], 1, doc)
                    .First(elem => elem.Text.Contains("Autofill" + "Semester"))
                    .Text = Section.Disciplines[discipline].Details.Keys.Any()
                        ? Section.Disciplines[discipline].Details.Keys
                        .Select(elem => elem.ToString())
                        .Aggregate((curr, next) => curr + ", " + next)
                        : "        ";



            /*
        Section.Disciplines[discipline].Details.Keys
            .Select(elem => elem.ToString())
            .Aggregate((curr, next) => curr + ", " + next);
            */
        }

        private void WriteCompetencies(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            var bookmarkElement = FindElementsByBookmark<Paragraph>(bookmarkMap["Competencies1"], 1, doc).First();
            var currentElement = bookmarkElement;

            foreach (var (name, classifier) in Section.CompetenceClassifiers)
            {
                if (!Section.DisciplineCompetencies[Section.Disciplines[discipline].Name].Any(competence => competence.StartsWith(name)))
                    continue;

                var element = new Paragraph
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "Default" },
                        Indentation = new Indentation { FirstLine = "709" },
                        Justification = new Justification { Val = JustificationValues.Both }
                    }
                };
                currentElement.InsertAfterSelf(element);
                currentElement = element;

                element.AppendChild(
                    new Run(new Text(classifier + ":"))
                    {
                        RunProperties = new RunProperties { Bold = new Bold() }
                    });

                var classifiedCompetencies = Section.Competencies
                    .Where(elem => Section.DisciplineCompetencies[Section.Disciplines[discipline].Name].Contains(elem.Key));

                Text text = default;

                foreach (var (_, competence) in classifiedCompetencies.Where(elem => elem.Key.StartsWith(name)))
                {
                    element = new Paragraph
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },
                            Indentation = new Indentation { FirstLine = "720" },
                            Justification = new Justification { Val = JustificationValues.Both }
                        }
                    };
                    currentElement.InsertAfterSelf(element);
                    currentElement = element;

                    var dot = competence.Name.IndexOf('.');
                    var rightText = competence.Name.Substring(0, dot + 1);

                    text = new Text(rightText) { Space = SpaceProcessingModeValues.Preserve };
                    element.AppendChild(
                        new Run(text)
                        {
                            RunProperties = new RunProperties
                            {
                                Bold = new Bold(),
                                RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                FontSize = new FontSize { Val = "24" },
                            }
                        });

                    rightText = competence.Name.Substring(dot + 1);
                    if (rightText[^1] == '.' || rightText[^1] == ';')
                        rightText = rightText.Substring(0, rightText.Length - 1);

                    text = new Text(rightText + ";") { Space = SpaceProcessingModeValues.Preserve };
                    element.AppendChild(
                        new Run(text)
                        {
                            RunProperties = new RunProperties
                            {
                                RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                FontSize = new FontSize { Val = "24" },
                            }
                        });
                }

                if (!(text is null))
                    text.Text = text.Text.Substring(0, text.Text.LastIndexOf(';')) + ".";

                //пустая строка
                element = new Paragraph();
                currentElement.InsertAfterSelf(element);
                currentElement = element;
            }

            bookmarkElement.Remove();
        }

        private void WriteYear(IDictionary<string, BookmarkStart> bookmarkMap, WordprocessingDocument doc) =>
            FindElementsByBookmark<Text>(bookmarkMap["Year1"], 1, doc)
                .First(elem => elem.Text.Contains("AutofillYear"))
                .Text = DateTime.Today.Year.ToString();

        private void WriteLaboriousnessTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            var rows = FindElementsByBookmark<Table>(bookmarkMap["LaboriousnessTable1"], 2, doc)
                .First()
                .Elements<TableRow>()
                .ToArray();


            //Добавляет в первую строку ячейку, которая далее будет сливаться с последующими
            rows.First().AppendChild(
                new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Restart }),
                    new Paragraph(
                        new Run(
                            new Text("Трудоемкость, академических часов")))));

            foreach (var (semester, details) in Section.Disciplines[discipline].Details)
            {
                var i = 0;

                rows.First().AppendChild(new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Continue }),
                    new Paragraph()));
                rows.Skip(1).First().AppendChild(GetTableCellByString($"{semester} семестр"));

                foreach (var row in rows.Skip(2))
                    row.AppendChild(GetTableCellByString(details[i++]));
            }

            WriteAtAllColumn();


            void WriteAtAllColumn()
            {
                var i = 0;
                var atAll = new DisciplineDetails();

                if (Section.Disciplines[discipline].Details.Count == 0)
                {
                    atAll = new DisciplineDetails
                    {
                        Semester = "0",
                        Monitoring = " , , , , , , , , ",
                        Contact = 0,
                        Lec = 0,
                        Lab = 0,
                        Pr = 0,
                        Ind = 0,
                        Control = 0,
                        Ze = 0
                    };
                }
                else
                {
                    atAll = new DisciplineDetails
                    {
                        Semester = Section.Disciplines[discipline].Details.Values.Select(details => details.Semester)
                            .Aggregate((current, next) => current + ", " + next),
                        Monitoring = Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring)
                            .Aggregate((current, next) => current + ", " + next),
                        Contact = Section.Disciplines[discipline].Details.Values.Sum(details => details.Contact),
                        Lec = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lec),
                        Lab = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lab),
                        Pr = Section.Disciplines[discipline].Details.Values.Sum(details => details.Pr),
                        Ind = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ind),
                        Control = Section.Disciplines[discipline].Details.Values.Sum(details => details.Control),
                        Ze = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ze)
                    };
                }
                rows.First().AppendChild(new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Continue }),
                    new Paragraph()));
                rows.Skip(1).First().AppendChild(GetTableCellByString("всего"));

                foreach (var row in rows.Skip(2))
                    row.AppendChild(GetTableCellByString(atAll[i++]));
            }
        }
        private void WriteLabVkrTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            var rows = FindElementsByBookmark<Table>(bookmarkMap["LabVkrTable1"], 2, doc)
                .First()
                .Elements<TableRow>()
                .ToArray();


            //Добавляет в первую строку ячейку, которая далее будет сливаться с последующими
            rows.First().AppendChild(
                new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Restart }),
                    new Paragraph(
                        new Run(
                            new Text("Трудоемкость, академических часов")))));

            foreach (var (semester, details) in Section.Disciplines[discipline].Details)
            {
                var i = 0;

                rows.First().AppendChild(new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Continue }),
                    new Paragraph()));
                rows.Skip(1).First().AppendChild(GetTableCellByString($"{semester} семестр"));

                foreach (var row in rows.Skip(2))
                {
                    if (i == 2)
                    {
                        row.AppendChild(GetTableCellByString(details[6]));
                        continue;
                    }
                    if (i == 3)
                    {
                        row.AppendChild(GetTableCellByString("Защита ВКР"));
                        continue;
                    }
                    row.AppendChild(GetTableCellByString(details[i++]));
                }
            }

            WriteAtAllColumn();


            void WriteAtAllColumn()
            {
                var i = 0;
                var atAll = new DisciplineDetails();

                if (Section.Disciplines[discipline].Details.Count == 0)
                {
                    atAll = new DisciplineDetails
                    {
                        Semester = "0",
                        Monitoring = " , , , , , , , , ",
                        Contact = 0,
                        Lec = 0,
                        Lab = 0,
                        Pr = 0,
                        Ind = 0,
                        Control = 0,
                        Ze = 0
                    };
                }
                else
                {
                    atAll = new DisciplineDetails
                    {
                        Semester = Section.Disciplines[discipline].Details.Values.Select(details => details.Semester)
                            .Aggregate((current, next) => current + ", " + next),
                        Monitoring = Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring)
                            .Aggregate((current, next) => current + ", " + next),
                        Contact = Section.Disciplines[discipline].Details.Values.Sum(details => details.Contact),
                        Lec = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lec),
                        Lab = Section.Disciplines[discipline].Details.Values.Sum(details => details.Lab),
                        Pr = Section.Disciplines[discipline].Details.Values.Sum(details => details.Pr),
                        Ind = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ind),
                        Control = Section.Disciplines[discipline].Details.Values.Sum(details => details.Control),
                        Ze = Section.Disciplines[discipline].Details.Values.Sum(details => details.Ze)
                    };
                }
                rows.First().AppendChild(new TableCell(
                    new TableCellProperties(
                        new HorizontalMerge { Val = MergedCellValues.Continue }),
                    new Paragraph()));
                rows.Skip(1).First().AppendChild(GetTableCellByString("всего"));

                foreach (var row in rows.Skip(2))
                {
                    if (i == 2)
                    {
                        row.AppendChild(GetTableCellByString(atAll[6]));
                        continue;
                    }
                    if (i == 3)
                    {
                        row.AppendChild(GetTableCellByString("Защита ВКР"));
                        continue;
                    }
                    row.AppendChild(GetTableCellByString(atAll[i++]));
                }
            }
        }
        //AutofillRequirements1
        //AutofillPracticleClassTable1
        //AutofillLaboratiesClassTable1

        private void WritePracticleClassTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            // Создание строки таблицы с настройкой стилей для ячеек
            var row = new TableRow(
                GetStyledTableCell(""),
                GetStyledTableCell(""),
                GetStyledTableCell("Итого:", false, "Times New Roman", 12, JustificationValues.Left),
                GetStyledTableCell(Section.Disciplines[discipline].Details.Values.Sum(details => details.Pr).ToString(), false, "Times New Roman", 12, JustificationValues.Center),
                GetStyledTableCell(""));

            // Находим таблицу и добавляем строку
            FindElementsByBookmark<Table>(bookmarkMap["PracticleClassTable1"], 2, doc)
                .First()
                .Append(row);
        }
        private void WriteLaboratiesClassTable(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            // Создание строки таблицы с настройкой стилей для ячеек
            var row = new TableRow(
                GetStyledTableCell(""),
                GetStyledTableCell(""),
                GetStyledTableCell("Итого:", false, "Times New Roman", 12, JustificationValues.Left),
                GetStyledTableCell(Section.Disciplines[discipline].Details.Values.Sum(details => details.Lab).ToString(), false, "Times New Roman", 12, JustificationValues.Center),
                GetStyledTableCell(""));

            // Находим таблицу и добавляем строку
            FindElementsByBookmark<Table>(bookmarkMap["LaboratiesClassTable1"], 2, doc)
                .First()
                .Append(row);
        }


        // Вспомогательный метод для создания стилизованной ячейки
        private TableCell GetStyledTableCell(
            string text,
            bool bold = false,
            string fontName = "Times New Roman",
            int fontSize = 12,
              JustificationValues? justification = null)
        {
            justification ??= JustificationValues.Left;
            var paragraph = new Paragraph
            {
                ParagraphProperties = new ParagraphProperties
                {
                    Justification = new Justification { Val = justification },
                }
            };

            var run = new Run();
            run.AppendChild(new Text(text));

            run.RunProperties = new RunProperties
            {
                Bold = bold ? new Bold() : null,
                RunFonts = new RunFonts { Ascii = fontName, HighAnsi = fontName },
                FontSize = new FontSize { Val = (fontSize * 2).ToString() },
            };

            paragraph.AppendChild(run);
            return new TableCell(paragraph);
        }

        //AutofillRequirementsLast
        //AutofillRequirementsNext

        private void WriteRequirements(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            if (!Section.Disciplines.ContainsKey(discipline) || !Section.DisciplineCompetencies.ContainsKey(Section.Disciplines[discipline].Name))
                return;

            var bookmarkElement1 = FindElementsByBookmark<Paragraph>(bookmarkMap["RequirementsLast1"], 1, doc).First();
            var bookmarkElement2 = FindElementsByBookmark<Paragraph>(bookmarkMap["RequirementsNext1"], 1, doc).First();
            var lastElement = bookmarkElement1;
            var nextElement = bookmarkElement2;
            var semesters = Section.Disciplines[discipline].Details.Keys;




            foreach (var elem in Section.Disciplines)
            {
                if (elem.Value.Details.Keys.Count < 1 || semesters.Count < 1)
                    continue;

                StringBuilder temp = new StringBuilder();
                temp.Append(elem.Value.Name + "(семестр ");
                foreach (var sem in elem.Value.Details.Keys)
                {
                    temp.Append(sem + " ,");
                }
                temp.Remove(temp.Length - 2, 2);
                temp.Append(')');

                if (elem.Value.Details.Keys.Max() < semesters.Min())
                {
                    var text = new Text(temp.ToString()) { Space = SpaceProcessingModeValues.Preserve };


                    var element = new Paragraph
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            ParagraphStyleId = new ParagraphStyleId { Val = "Default" },
                            // Одинарный межстрочный интервал: LineRule = Auto и Line = 240 примерно соответствуют одинарному интервалу.
                            SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },

                            // Отступ первой строки (например, 720 соответствует ~0.5 дюйма)
                            Indentation = new Indentation { FirstLine = "720" },
                            Justification = new Justification { Val = JustificationValues.Both }
                        }
                    };
                    lastElement.InsertAfterSelf(element);
                    lastElement = element;


                    element.AppendChild(
                             new Run(text)
                             {
                                 RunProperties = new RunProperties
                                 {
                                     Bold = new Bold(),
                                     RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                     FontSize = new FontSize { Val = "24" }, // 12 пт (24 полукегля)

                                     // Выделение текста жёлтым цветом
                                     Highlight = new Highlight { Val = HighlightColorValues.Yellow }
                                 }
                             }
                     );
                }
                else
                {

                    var text = new Text(temp.ToString()) { Space = SpaceProcessingModeValues.Preserve };

                    var element2 = new Paragraph
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            ParagraphStyleId = new ParagraphStyleId { Val = "Default" },
                            // Одинарный межстрочный интервал: LineRule = Auto и Line = 240 примерно соответствуют одинарному интервалу.
                            SpacingBetweenLines = new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto },

                            // Отступ первой строки (например, 720 соответствует ~0.5 дюйма)
                            Indentation = new Indentation { FirstLine = "720" },
                            Justification = new Justification { Val = JustificationValues.Both }
                        }
                    };
                    nextElement.InsertAfterSelf(element2);
                    nextElement = element2;
                    element2.AppendChild(
                        new Run(text)
                        {
                            RunProperties = new RunProperties
                            {
                                Bold = new Bold(),
                                RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                FontSize = new FontSize { Val = "24" }, // 12 пт (24 полукегля)

                                // Выделение текста жёлтым цветом
                                Highlight = new Highlight { Val = HighlightColorValues.Yellow }
                            }
                        }
                    );
                }
            }

            bookmarkElement1.Remove();
            bookmarkElement2.Remove();
        }

        private void CheckKurs(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc, 
            string templatePath, string dwpDir, Employee employes)
        {
            //isSave = false;
            HashSet<string> set = new HashSet<string>();
            var temp = Section.Disciplines.Where(x => x.Value.Name.Contains(Section.Disciplines[discipline].Name)
                    && x.Value.Name.Contains("Проект по", StringComparison.OrdinalIgnoreCase)).
                    Select(x => x.Key).FirstOrDefault();


            foreach (var mon in Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring))
            {
                foreach (var item in mon.Split(' '))
                {

                    if (item == "КР")
                    {
                        set.Add("КР");
                    }
                    if (item == "КП")
                    {
                        set.Add("КП");
                    }
                }
            }
            if (temp != null)
            {
                foreach (var mon in Section.Disciplines[temp].Details.Values.Select(details => details.Monitoring))
                {
                    foreach (var item in mon.Split(' '))
                    {

                        if (item == "КР")
                        {
                            set.Add("КР");
                        }
                        if (item == "КП")
                        {
                            set.Add("КП");
                        }
                    }
                }
            }

            if (set.Contains("КР"))
            {
                FindElementsByBookmark<Text>(bookmarkMap["Kurs1"], 1, doc)
                .First(elem => elem.Text.Contains("AutofillKurs"))
                .Text = "учебным планом предусмотренна курсовая работа. Пожалуйста заполните этот пункт.";
            }
            else if (set.Contains("КП"))
            {
                FindElementsByBookmark<Text>(bookmarkMap["Kurs1"], 1, doc)
                .First(elem => elem.Text.Contains("AutofillKurs"))
                .Text = "учебным планом предусмотрен курсовой проект. Пожалуйста заполните этот пункт.";
            }
            else
            {
                FindElementsByBookmark<Text>(bookmarkMap["Kurs1"], 1, doc)
                .First(elem => elem.Text.Contains("AutofillKurs"))
                .Text = "учебным планом курсовой проект не предусмотрен.";
            }

        } 
    }
}