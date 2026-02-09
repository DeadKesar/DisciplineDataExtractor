using System.IO;
using System.Text.RegularExpressions;

namespace DisciplineWorkProgram.Word
{
    public static class RegexPatterns
    {
        //Некорректные символы в пути файла
        public static readonly Regex InvalidChars =
            new Regex($"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars()))}]");

    }
}
