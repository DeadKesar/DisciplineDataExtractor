using System.Text.RegularExpressions;

namespace DisciplineDataExtractor.Models.Sections
{
    public class RegexPatterns
    {
        //Наименование компетенции тут довольно сложный случай так как приходится учитывать много вариантов и ошибок.
        //старый регекс на всякий: ^(УК-\d{1,2}.?(! )?(?!\d))|^(ОПК-\d{1,2}.?(! )?(?!\d))|^(ПК-\d{1,2}.?(! )?(?!\d))
        public static readonly Regex CompetenceName = new Regex(@"^(УК-\s*[\dЗ]{1,2}.?(! )?(?![\dЗ]))|^(ОПК-\s*[\dЗ]{1,2}.?(! )?(?![\dЗ]))|^(ПК-\s*[\dЗ]{1,2}.?(! )?(?![\dЗ]))");

        public static readonly Regex CompetenceName2 = new Regex(@"^(УК-\s*[\dЗ]+(\.\d+)*|ОПК-\s*[\dЗ]+(\.\d+)*|ПК-\s*[\dЗ]+(\.[\dЗ]+)*)\b");
        //Любая строка, содержащая в себе информацию о компетенции
        public static readonly Regex Competence = new Regex(@"^(УК-\s*)|^(ОПК-\s*)|^(ПК-\s*)");

        public static readonly Regex VrongCompetence = new Regex(@"^(УК-\s* З||(\dЗ))|^(ОПК-\s* З||(\dЗ))|^(ПК-\s* З||(\dЗ))");
        //Поиск в строке ...?
        public static readonly Regex WayNameSection = new Regex("(?<=\").*(?=\")");
        //Поиск числа в строке
        public static readonly Regex DigitInString = new Regex(@"\d");
    }
}
