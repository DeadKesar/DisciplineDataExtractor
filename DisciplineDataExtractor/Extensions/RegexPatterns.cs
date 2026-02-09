using System.Text.RegularExpressions;

namespace DisciplineDataExtractor.Extensions
{
    public static class RegexPatterns
    {
        //Повторяющиеся пробелы
        public static readonly Regex MultipleSpaces = new Regex("[ ]{2,}");
    }
}
