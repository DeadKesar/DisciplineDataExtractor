using System.Collections.Generic;

namespace DisciplineDataExtractor.Models.Sections
{
    public class Sections
    {
        public IDictionary<string, Section> Type { get; set; } = new Dictionary<string, Section>();
    }
}
