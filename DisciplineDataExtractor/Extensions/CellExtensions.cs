using ClosedXML.Excel;

namespace DisciplineDataExtractor.Extensions
{
    public static class CellExtensions
    {
        public static int GetInt(this IXLCell cell)
        {
            try
            {
                return cell.GetValue<int>();
            }
            catch
            {
                return 0;
            }
        }
    }
}
