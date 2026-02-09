using ClosedXML.Excel;
using DisciplineDataExtractor.Models;
using DisciplineDataExtractor.Models.Sections;
using System.Text.Json;
using System.Collections.Generic; // обязательно

namespace DisciplineDataExtractor
{
    public static class DataExtractor
    {
        public static string ExtractJson(
    Stream planStream,
    Stream compListStream,
    Stream compMatrixStream,
    Stream? dolznostiStream = null)
        {
            // Перемотка
            planStream.Seek(0, SeekOrigin.Begin);
            compListStream.Seek(0, SeekOrigin.Begin);
            compMatrixStream.Seek(0, SeekOrigin.Begin);
            dolznostiStream?.Seek(0, SeekOrigin.Begin);

            // Сотрудники (эффективно: один workbook)
            Employee? employees = null;
            if (dolznostiStream != null)
            {
                using var dolznostiWorkbook = new XLWorkbook(dolznostiStream);
                employees = new Employee(dolznostiWorkbook);
            }

            // Парсинг
            var section = new Section(compListStream, compMatrixStream);
            section.LoadDataFromPlan(planStream);
            section.LoadCompetenciesData();

            // Результат (с TryGetValue для компетенций)
            var disciplinesDict = new Dictionary<string, object>();
            foreach (var d in section.Disciplines)
            {
                section.DisciplineCompetencies.TryGetValue(d.Value.Name, out var compList);
                disciplinesDict[d.Key] = new
                {
                    d.Value.Ind,
                    d.Value.Name,
                    d.Value.Department,
                    Details = d.Value.Details,
                    Competencies = compList ?? new List<string>()
                };
            }

            var result = new
            {
                SectionInfo = section.SectionDictionary,
                Disciplines = disciplinesDict,
                Competencies = section.Competencies,
                Employees = employees?.Employees ?? new Dictionary<string, IDictionary<string, string>>()
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.Create(System.Text.Unicode.UnicodeRanges.All)
            };

            return JsonSerializer.Serialize(result, options);
        }
    }
}