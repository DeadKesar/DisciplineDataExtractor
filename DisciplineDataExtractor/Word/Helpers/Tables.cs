using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace DisciplineWorkProgram.Word.Helpers
{
    public static class Tables
    {
        public static IEnumerable<TableCell> GetTablesCells(WordprocessingDocument doc) =>
            doc.MainDocumentPart.Document.Body.Descendants<TableCell>();

        public static IEnumerable<Table> GetTables(WordprocessingDocument document) =>
            document.MainDocumentPart.Document.Body.Descendants<Table>();

        public static IEnumerable<string> GetHeaders(Table table) =>
            table
                .Descendants<TableRow>().First()
                .Descendants<Paragraph>().Select(p => p.InnerText);

        //На каждую строку по параграфу
        public static TableCell GetTableCellByStrings(IEnumerable<string> values)
        {
            var cell = new TableCell();

            foreach (var value in values)
                cell.AppendChild(new Paragraph(new Run(new Text(value))));

            return cell;
        }

        public static IEnumerable<TableCell> GetTableCellsByStrings(params string[] values) =>
            values.Select(value => new TableCell(new Paragraph(new Run(new Text(value)))));


        public static TableCell GetTableCellByString(string value) =>
            new TableCell(new Paragraph(new Run(new Text(value))));

    }
}
