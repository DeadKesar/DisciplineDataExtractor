using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using DisciplineDataExtractor.Extensions; // для RemoveMultipleSpaces
using DisciplineDataExtractor.Models.Sections; // для RegexPatterns

namespace DisciplineDataExtractor.Models.Sections.Helpers
{
    public static class Competencies
    {
        public static IEnumerable<string> ParseCompetencies(WordprocessingDocument document)
        {
            var competencies = new List<string>();

            foreach (var table in document.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>())
            {
                foreach (var cell in table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                {
                    if (!cell.Descendants<Text>().Any(t => RegexPatterns.Competence.IsMatch(t.Text)))
                        continue;

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

                    if (!string.IsNullOrEmpty(tmp))
                        competencies.Add(tmp.RemoveMultipleSpaces());
                }
            }

            return competencies;
        }
    }
}