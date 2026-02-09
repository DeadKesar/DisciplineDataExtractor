using System.Collections.Generic;

namespace DisciplineWorkProgram.Models.Sections
{
    public class Sections
    {
        public IDictionary<string, Section> Type { get; set; } = new Dictionary<string, Section>();
    }
}
