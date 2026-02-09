using System.Collections.Generic;

namespace DisciplineWorkProgram.Models
{
    public class Competence
    {
        public string Name { get; set; }

        public IList<string> Competencies { get; set; } = new List<string>();
    }
}
