#nullable disable

using System.Xml;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.RichText;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Worksheets;
using static Aid.Xlsx.SpreadsheetML.OpenXmlConst;

namespace Aid.Xlsx.SpreadsheetML
{
    internal class CommentPartWriter
    {
        internal static void GenerateWorksheetCommentsPartContent(Stream partStream,
            XlsxWorksheet xlWorksheet)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = XlsxHelper.NoBomUTF8
            };
            using var xml = XmlWriter.Create(partStream, settings);

            var commentCells = new List<XlsxCell>();
            var authorsDict = new Dictionary<String, Int32>();
            xml.WriteStartElement("x", "comments", Main2006SsNs);
            foreach (var c in xlWorksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
            {
                var authorName = c.GetComment().Author;

                if (!authorsDict.TryGetValue(authorName, out var authorId))
                {
                    authorId = authorsDict.Count;
                    authorsDict.Add(authorName, authorId);
                }

                commentCells.Add(c);
            }

            xml.WriteStartElement("authors", Main2006SsNs);
            foreach (var author in authorsDict)
                xml.WriteElementString("author", Main2006SsNs, author.Key);

            xml.WriteEndElement(); // authors

            var refBuffer = new char[10];
            xml.WriteStartElement("commentList", Main2006SsNs);
            foreach (var commentCell in commentCells)
            {
                var comment = commentCell.GetComment();
                xml.WriteStartElement("comment", Main2006SsNs);

                var refLen = commentCell.Point.Format(refBuffer);
                xml.WriteStartAttribute("ref");
                xml.WriteRaw(refBuffer, 0, refLen);
                xml.WriteEndAttribute(); // ref

                var authorId = authorsDict[comment.Author];
                xml.WriteAttribute("authorId", authorId);

                // Excel specifies @guid is optional if the workbook is not shared
                // Excel ignores the shapeId attribute.

                xml.WriteStartElement("text", Main2006SsNs);
                var richText = XlsxImmutableRichText.Create(comment);
                foreach (var run in richText.Runs)
                    TextSerializer.WriteRun(xml, richText, run);

                xml.WriteEndElement(); // text
                xml.WriteEndElement(); // comment
            }

            xml.WriteEndElement(); // commentList
            xml.WriteEndElement(); // comments

            xml.Close();
        }
    }
}
