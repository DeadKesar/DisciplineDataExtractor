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
    public class Fos
    {
        public Fos(Section section)
        {
            Section = section;
        }

        private Section Section { get; }

        //Должно обрабатывать только 1 дисциплину, чтобы "масштабировать" без доп. кода
        public void MakeFos(string templatePath, string fosDir, string discipline, Employee employes)
        {

            using var doc = WordprocessingDocument.CreateFromTemplate(templatePath, true);
            var bookmarkMap = GetBookmarks(doc, "Autofill");

            
            WriteSectionData(bookmarkMap, doc);
            WriteDisciplineData(bookmarkMap, discipline, doc);
            WriteEmploesData(bookmarkMap, discipline, doc, employes);
            //WriteRequirements(bookmarkMap, discipline, doc);
            WriteCompetenciesTable(bookmarkMap, discipline, doc); //заполняет табличку компетенций
            //WriteDisciplinePartitionTable(bookmarkMap, discipline, doc);
            //WritePracticleClassTable(bookmarkMap, discipline, doc);
            //WriteSemesters(bookmarkMap, discipline, doc);
            //WriteCompetencies(bookmarkMap, discipline, doc);//записываем компетенции в самом начале
            //WriteYear(bookmarkMap, doc);
            // Не реализовано занесение данных по дисциплине
            //WriteLaboriousnessTable(bookmarkMap, discipline, doc);
            //WriteLaboratiesClassTable(bookmarkMap, discipline, doc);
            CheckExam(bookmarkMap, discipline, doc);
            SaveDoc(doc, fosDir, Section.Disciplines[discipline].Name);
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

        private void CheckExam(IDictionary<string, BookmarkStart> bookmarkMap, string discipline, WordprocessingDocument doc)
        {
            HashSet<string> set = new HashSet<string>();
            foreach (var mon in Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring))
            {
                foreach (var item in mon.Split(' '))
                {
                    if (item == "За")
                    {
                        set.Add("Зачёт");
                    }
                    if (item == "Эк")
                    {

                        set.Add("Экзамен");
                    }
                    if (item == "ЗаО")
                    {
                        set.Add("Зачёт с оценкой");
                    }
                    if (item == "КР")
                    {
                        continue;
                    }
                }
            }

            if (!set.Contains("Экзамен") && !set.Contains("Зачёт с оценкой"))
            {

                var body = doc.MainDocumentPart.Document.Body;

                // Находим начало и конец закладки
                var bookmarkStart = body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == "AutofillExaming1");
                var bookmarkEnd = body.Descendants<BookmarkEnd>().FirstOrDefault(be => be.Id == bookmarkStart?.Id);

                if (bookmarkStart == null || bookmarkEnd == null)
                {
                    throw new System.Exception($"Закладка AutofillExaming1 не найдена или не имеет конца.");
                }

                // Собираем все элементы между BookmarkStart и BookmarkEnd
                var elementsToRemove = body.Descendants()
                    .SkipWhile(e => e != bookmarkStart) // Пропустить элементы до начала закладки
                    .Skip(1) // Пропустить сам BookmarkStart
                    .TakeWhile(e => e != bookmarkEnd) // Взять все элементы до BookmarkEnd
                    .ToList();

                // Удаляем найденные элементы
                foreach (var element in elementsToRemove)
                {
                    element.Remove();
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
                //Надежда на то, что понадобится нумерация в рамках только одноразрядного числа, это о том что количество
                //закладок с одинаковым именем не может быть больше 10
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
                if (actualKey == "Year")
                {
                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("AutofillYear"))
                        .Text = DateTime.Today.Year.ToString();
                }
                if (actualKey == "PassDestination")
                {
                    String name;
                    if (Section.Disciplines[discipline].Props["Name"].Contains("практик"))
                    {
                        name = "фонда оценочных средств по практике: " + Section.Disciplines[discipline].Props["Name"];
                    }
                    else if (Section.Disciplines[discipline].Props["Name"].Contains("выпускной квалификационной работы"))
                    {
                        name = "фонда оценочных средств по ГИА: " + Section.Disciplines[discipline].Props["Name"];
                    }
                    else
                    {
                        name = "фонда оценочных средств по дисциплине: " + Section.Disciplines[discipline].Props["Name"];
                    }
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = name;
                    continue;
                }
                //"AutofillMonitoring" = "За"
                if (actualKey == "Monitoring")
                {
                    StringBuilder temp = new StringBuilder();
                    HashSet<string> set = new HashSet<string>();
                    foreach (var mon in Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring)) {
                        foreach (var item in mon.Split(' '))
                        {
                            if (item == "За")
                            {
                                set.Add("Зачёта");
                            }
                            if (item == "Эк")
                            {
 
                                set.Add("Экзамена");
                            }
                            if (item == "ЗаО")
                            {
                                set.Add("Зачёта с оценкой");
                            }
                            if (item == "КР")
                            {
                                continue;
                            }
                        }
                    }
                    foreach(var mon in set)
                    {
                        if (temp.Length > 0)
                            temp.Append(", ");

                        temp.Append(mon);
                    }

                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = temp.ToString();
                    continue;
                }
                if (actualKey == "MonitoringU")
                {
                    StringBuilder temp = new StringBuilder();
                    HashSet<string> set = new HashSet<string>();
                    foreach (var mon in Section.Disciplines[discipline].Details.Values.Select(details => details.Monitoring))
                    {
                        foreach (var item in mon.Split(' '))
                        {
                            if (item == "За")
                            {
                                set.Add("Зачёту");
                            }
                            if (item == "Эк")
                            {

                                set.Add("Экзамену");
                            }
                            if (item == "ЗаО")
                            {
                                set.Add("Зачёту с оценкой");
                            }
                            if (item == "КР")
                            {
                                continue;
                            }
                        }
                    }
                    foreach (var mon in set)
                    {
                        if (temp.Length > 0)
                            temp.Append(", ");

                        temp.Append(mon);
                    }

                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = temp.ToString();
                    continue;
                }

                if (actualKey == "EducationLvl")
                {
                    //Специальность / направление 
                    String edLVL;
                    if (Section.SectionDictionary["EducationLevel"] == "Специалитет")
                    {
                        edLVL = "Специальность";
                    }
                    else
                        edLVL = "Направление";
                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = edLVL.ToString();
                    continue;
                }
                if (actualKey == "Curs")
                {
                    StringBuilder temp = new StringBuilder();

                    foreach (var sem in Section.Disciplines[discipline].Details.Values.Select(details => details.Semester))
                    {
                        int num = (int)((int.Parse(sem) - 1) / 2) + 1;
                        temp.Append(num.ToString() + " курса(" + sem + " семестра)");
                        temp.Append(" ");
                    }

                    FindElementsByBookmark<Text>(bookmark, 1, doc)
                        .First(elem => elem.Text.Contains("Autofill" + actualKey))
                        .Text = temp.ToString();
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
                    case "ProOD":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Проректор по образовательной деятельности и молодежной политике"]["nameForDoc"];
                        continue;
                    case "ProODName":
                        FindElementsByBookmark<Text>(bookmark, 1, doc)
                            .First(elem => elem.Text.Contains("Autofill" + actualKey))
                            .Text = employes.Employees["Проректор по образовательной деятельности и молодежной политике"]["FIO"];
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


                var firstCell = new TableCell(
                new TableCellProperties(
                    new VerticalMerge { Val = MergedCellValues.Restart }
                    ),
                    new Paragraph(new Run(new Text(Section.Competencies[competence].Name)))
                );

                var relatedCompetencies = Section.Competencies
                    .Where(kvp => kvp.Key.StartsWith(competence + "."))
                    .ToList();
                //var cell = new TableCell();

                //var paragraph = new Paragraph();
                bool isFirst = true;

                if (relatedCompetencies.Count == 0)
                {
                    var emptyRow = new TableRow();
                    emptyRow.Append(firstCell);
                    emptyRow.Append(new TableCell(new Paragraph(new Run(new Text("")))));
                    emptyRow.Append(new TableCell(new Paragraph(new Run(new Text("")))));
                    emptyRow.Append(new TableCell(new Paragraph(new Run(new Text("")))));
                    table.AppendChild(emptyRow);
                    continue;
                }

                foreach (var s in relatedCompetencies)
                {
                    var row = new TableRow();
                    if (isFirst)
                    {
                        row.Append(firstCell);
                        isFirst = false;
                    }
                    else
                    {
                        row.Append(new TableCell(
                            new TableCellProperties(
                                new VerticalMerge { Val = MergedCellValues.Continue }
                            ), new Paragraph()
                        ));
                    }
                    var paragraph = new Paragraph(new Run(new Text(s.Value.Name)));
                    var tableCell = new TableCell(paragraph);
                    row.AppendChild(tableCell);
                    row.AppendChild(new TableCell(new Paragraph(new Run(new Text("")))));
                    row.AppendChild(new TableCell(new Paragraph(new Run(new Text("")))));

                    table.AppendChild(row);

                }

            }
        }

    }
}
