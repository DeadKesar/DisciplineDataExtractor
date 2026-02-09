using DisciplineWorkProgram.Extensions;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using static DisciplineDataExtractor.Word.Helpers.Tables;

namespace DisciplineWorkProgram.Models.Sections.Helpers
{
    public class Competencies
    {
        public static IEnumerable<string> ParseCompetencies(WordprocessingDocument document)
        {
            var competencies = new List<string>();
            foreach (var cell in GetTablesCells(document)
                .Where(cell => cell.Descendants<Text>().Any(text => RegexPatterns.Competence.IsMatch(text.Text))))
            {
                var tmp = string.Empty;
                foreach (var text in cell.Descendants<Text>())
                {
                    if (RegexPatterns.Competence.IsMatch(text.Text) && !string.IsNullOrEmpty(tmp))
                    {
                        competencies.Add(tmp.RemoveMultipleSpaces());
                        tmp = string.Empty;
                    }
                    tmp += text.Text;
                }
                if (!string.IsNullOrEmpty(tmp)) competencies.Add(tmp.RemoveMultipleSpaces());
            }

            return competencies;
        }
    }
}
