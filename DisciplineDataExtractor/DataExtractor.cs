using ClosedXML.Excel;
using DisciplineDataExtractor.Excel;
using DisciplineDataExtractor.Models;
using DisciplineDataExtractor.Models.Sections;
using System.Collections.Generic; // обязательно
using System.Text.Json;

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
            // Перемотка всех потоков
            planStream.Seek(0, SeekOrigin.Begin);
            compListStream.Seek(0, SeekOrigin.Begin);
            compMatrixStream.Seek(0, SeekOrigin.Begin);
            dolznostiStream?.Seek(0, SeekOrigin.Begin);

            // Автоматическая конвертация .xls → .xlsx (если нужно)
            try
            {
                // Пытаемся открыть как .xlsx (ClosedXML бросит исключение, если .xls)
                using var testWorkbook = new XLWorkbook(planStream);
                // Если успешно — это .xlsx, просто перематываем для дальнейшего чтения
                planStream.Seek(0, SeekOrigin.Begin);
            }
            catch
            {
                // Если не .xlsx — конвертируем в .xlsx-Stream
                planStream.Seek(0, SeekOrigin.Begin); // обязательно перемотать после неудачной попытки
                planStream = Converter.ConvertToXlsxStream(planStream);
            }

            // Сотрудники
            Employee? employees = null;
            if (dolznostiStream != null)
            {
                using var dolznostiWorkbook = new XLWorkbook(dolznostiStream);
                employees = new Employee(dolznostiWorkbook);
            }

            // Парсинг
            var section = new Section(compListStream, compMatrixStream);
            section.LoadDataFromPlan(planStream); // теперь поток точно с начала
            section.LoadCompetenciesData();

            // Результат
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