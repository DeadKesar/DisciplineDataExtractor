using ClosedXML.Excel;
using System.Text.Json.Serialization;
using static DisciplineDataExtractor.Models.Sections.Helpers.ExcelFinder;

namespace DisciplineDataExtractor.Models
{
    public class Employee : HierarchicalCheckableElement
    {
        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Enumerable.Empty<HierarchicalCheckableElement>();

        [JsonIgnore]
        public IDictionary<string, IDictionary<string, string>> Employees { get; } = new Dictionary<string, IDictionary<string, string>>();

        public Employee(Stream dolznostiStream)
        {
            dolznostiStream.Seek(0, SeekOrigin.Begin);
            using var workbook = new XLWorkbook(dolznostiStream);
            LoadEmployees(workbook);
        }

        // перегрузка под XLWorkbook
        public Employee(IXLWorkbook workbook)
        {
            LoadEmployees(workbook);
        }
        private void LoadEmployees(IXLWorkbook workbook)
        {
            var worksheet = workbook.Worksheets
                .SingleOrDefault(sheet => sheet.Name.StartsWith("Сотрудники"));
            if (worksheet == null)
                throw new InvalidOperationException("Не найден лист 'Сотрудники'.");

            foreach (var row in worksheet.RowsUsed().Where(row => int.TryParse(row.Cell(FindColumn(worksheet, "номер")).GetString(), out _)))
            {
                var emp = row.Cell("B").GetString();
                var employeeData = new Dictionary<string, string>
                {
                    ["nameForDoc"] = row.Cell("C").GetString(),
                    ["position"] = row.Cell("D").GetString(),
                    ["FIO"] = row.Cell("E").GetString(),
                    ["institut"] = row.Cell("F").GetString(),
                };

                // Добавляем сотрудника в словарь
                Employees[emp] = employeeData;
            }
        }
        public Employee(string path) : this(new XLWorkbook(path)) { }
    }
}
