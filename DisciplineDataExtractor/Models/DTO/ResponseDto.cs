using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DisciplineDataExtractor.Models.Dto
{
    /// <summary>
    /// Основной ответ микросервиса — все извлечённые данные из учебного плана
    /// </summary>
    public class ExtractResponse
    {
        /// <summary>
        /// Общая информация о направлении подготовки
        /// </summary>
        public SectionInfo SectionInfo { get; set; } = new();

        /// <summary>
        /// Все дисциплины (ключ — индекс дисциплины)
        /// </summary>
        public Dictionary<string, DisciplineDto> Disciplines { get; set; } = new();

        /// <summary>
        /// Полный справочник компетенций
        /// </summary>
        public Dictionary<string, CompetenceDto> Competencies { get; set; } = new();

        /// <summary>
        /// Справочник сотрудников и кафедр
        /// </summary>
        public Dictionary<string, EmployeeDto> Employees { get; set; } = new();
    }

    public class SectionInfo
    {
        public string EducationLevel { get; set; } = string.Empty;
        public string WayCode { get; set; } = string.Empty;
        public string EducationForm { get; set; } = string.Empty;
        public string WayName { get; set; } = string.Empty;
        public string WaySection { get; set; } = string.Empty;
    }

    public class DisciplineDto
    {
        public string Ind { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;

        /// <summary>
        /// Распределение по семестрам (ключ — номер семестра)
        /// </summary>
        public Dictionary<int, DisciplineDetailsDto> Details { get; set; } = new();

        /// <summary>
        /// Список кодов компетенций, привязанных к дисциплине
        /// </summary>
        public List<string> Competencies { get; set; } = new();
    }

    public class DisciplineDetailsDto
    {
        public string Monitoring { get; set; } = string.Empty;
        public int Contact { get; set; }
        public int Lec { get; set; }
        public int Lab { get; set; }
        public int Pr { get; set; }
        public int Ind { get; set; }
        public int Control { get; set; }
        public int Ze { get; set; }
        public string Semester { get; set; } = string.Empty;
    }

    public class CompetenceDto
    {
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Подкомпетенции (если есть)
        /// </summary>
        public List<string> Competencies { get; set; } = new();
    }

    public class EmployeeDto
    {
        public string NameForDoc { get; set; } = string.Empty;
        public string Position { get; set; } = string.Empty;
        public string FIO { get; set; } = string.Empty;
        public string Institut { get; set; } = string.Empty;
    }
}