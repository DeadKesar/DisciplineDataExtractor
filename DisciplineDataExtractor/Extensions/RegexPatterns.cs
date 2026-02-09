using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Extensions
{
    public static class RegexPatterns
    {
        //Повторяющиеся пробелы
        public static readonly Regex MultipleSpaces = new Regex("[ ]{2,}");
    }
}
