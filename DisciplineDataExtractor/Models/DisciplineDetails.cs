namespace DisciplineWorkProgram.Models
{
    public class DisciplineDetails
    {
        public string Monitoring { get; set; }
        public int Contact { get; set; }
        public int Lec { get; set; }
        public int Lab { get; set; }
        public int Pr { get; set; }
        public int Ind { get; set; }
        public int Control { get; set; }
        public int Ze { get; set; }
        public string Semester { get; set; }

        public bool IsHollow => string.IsNullOrWhiteSpace(Monitoring) &&
                                string.IsNullOrWhiteSpace(Semester) &&
                              Contact == 0 &&
                              Lec == 0 &&
                              Lab == 0 &&
                              Pr == 0 &&
                              Ind == 0 &&
                              Control == 0 &&
                              Ze == 0;

        /// <summary>
        /// Очень опасная штука, применяется только для записи деталей в таблицу.
        /// 0 - Общая трудоёмкость,
        /// 1 - Контактная работа,
        /// 2 - Лекции,
        /// 3 - Практические занятия,
        /// 4 - Лабы,
        /// 5 - Промеж. контроль,
        /// 6 - Самост. раб.,
        /// 7 - заглушка,
        /// 8 - Вид контроля
        /// 9 - семестр
        /// </summary>
        /// <param name="i"></param>
        public string this[int i] => i switch
        {
            0 => (Contact + Ind + Control).ToString(),
            1 => Contact.ToString(),
            2 => Lec.ToString(),
            3 => Pr.ToString(),
            4 => Lab.ToString(),
            5 => Control.ToString(),
            6 => Ind.ToString(),
            7 => string.Empty,
            8 => Monitoring,
            9 => Semester,
            _ => string.Empty
        };

    }
}
