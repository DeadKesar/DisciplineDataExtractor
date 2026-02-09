using System.IO;

namespace DisciplineDataExtractor.Excel
{
    internal class NpoiMemoryStream : MemoryStream
    {
        public bool AllowClose { get; set; } = true;

        public override void Close()
        {
            if (AllowClose) base.Close();
        }
    }
}
