using System.Linq;

namespace DisciplineWorkProgram.Extensions
{
    public static class StringExtensions
    {
        public static string RemoveMultipleSpaces(this string str) =>
            RegexPatterns.MultipleSpaces.Replace(str, " ").Trim();

        public static bool ContainsAny(this string str, params string[] values) =>
            values.Any(str.Contains);
    }
}
