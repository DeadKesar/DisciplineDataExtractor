using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace DisciplineDataExtractor.Models
{
    public class Discipline : HierarchicalCheckableElement
    {
        protected override IEnumerable<HierarchicalCheckableElement> GetNodes() => Enumerable.Empty<HierarchicalCheckableElement>();

        //Название свойства, значение
        [JsonIgnore]
        public IDictionary<string, string> Props { get; } = new Dictionary<string, string>();
        //Семестр, детали работ
        public IDictionary<int, DisciplineDetails> Details { get; } = new Dictionary<int, DisciplineDetails>();

        public string Ind
        {
            get => Props["Discipline"];
            set => Props["Discipline"] = value;

            //get => Props["Ind"];
            //set => Props["Ind"] = value;
            //get => Convert.ToInt32(Props["Ind"]);
            //set => Props["Ind"] = value.ToString();
        }
        public string Name
        {
            get => Props["Name"];
            set => Props["Name"] = value;

            //get => Props["Discipline"];
            //set => Props["Discipline"] = value;
        }

        public string Department
        {
            get => Props["Department"];
            set => Props["Department"] = value;
        }

        public int Exam
        {
            get => Convert.ToInt32(Props["Exam"]);
            set => Props["Exam"] = value.ToString();
        }
        public int Credit
        {
            get => Convert.ToInt32(Props["Credit"]);
            set => Props["Credit"] = value.ToString();
        }

        public int CreditWithRating
        {
            get => Convert.ToInt32(Props["CreditWithRating"]);
            set => Props["CreditWithRating"] = value.ToString();
        }

        public int Kp
        {
            get => Convert.ToInt32(Props["KP"]);
            set => Props["KP"] = value.ToString();
        }
        public int Kr
        {
            get => Convert.ToInt32(Props["KR"]);
            set => Props["KR"] = value.ToString();
        }

        public int Fact
        {
            get => Convert.ToInt32(Props["Fact"]);
            set => Props["Fact"] = value.ToString();
        }

        public int ByPlan
        {
            get => Convert.ToInt32(Props["ByPlan"]);
            set => Props["ByPlan"] = value.ToString();
        }

        public int ContactHours
        {
            get => Convert.ToInt32(Props["ContactHours"]);
            set => Props["ContactHours"] = value.ToString();
        }

        public int Lec
        {
            get => Convert.ToInt32(Props["Lec"]);
            set => Props["Lec"] = value.ToString();
        }

        public int Lab
        {
            get => Convert.ToInt32(Props["Lab"]);
            set => Props["Lab"] = value.ToString();
        }

        public int Pr
        {
            get => Convert.ToInt32(Props["Pr"]);
            set => Props["Pr"] = value.ToString();
        }

        public int Control
        {
            get => Convert.ToInt32(Props["Control"]);
            set => Props["Control"] = value.ToString();
        }

        public int ZeAtAll
        {
            get => Convert.ToInt32(Props["ZeAtAll"]);
            set => Props["ZeAtAll"] = value.ToString();
        }
    }
}
