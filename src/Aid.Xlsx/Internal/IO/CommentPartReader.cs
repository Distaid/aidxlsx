using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using Aid.Xlsx.Opc;

namespace Aid.Xlsx.Internal.IO;

internal static class CommentPartReader
{
    internal static void Load(Stream input, XLWorksheet worksheet, Func<int, int?> getShapeId, Action<int, XLComment> applyShape)
    {
        using var xml = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
        xml.MoveToContent();
        var ns = xml.NamespaceURI;
        using var tree = new XmlTreeReader(xml, XmlToEnumMapper.Instance, worksheet.Workbook.StrictAttributeParsing, leaveOpen: true);
        if (!tree.TryOpen("comments", ns))
            throw PartStructureException.ExpectedElementNotFound("comments", tree);

        var authors = new List<string>();
        if (!tree.TryOpen("authors", ns))
            throw PartStructureException.ExpectedElementNotFound("authors", tree);
        while (tree.TryOpen("author", ns))
        {
            authors.Add(tree.GetContent());
            tree.Close("author", ns);
        }

        tree.Close("authors", ns);
        if (!tree.TryOpen("commentList", ns))
            throw PartStructureException.ExpectedElementNotFound("commentList", tree);

        var textReader = new RstReader(tree, worksheet.Workbook.Styles, ns);
        var index = 0;
        while (tree.TryOpen("comment", ns))
        {
            var address = tree.GetString("ref");
            var author = tree.GetUInt("authorId");
            if (author >= authors.Count)
                throw new InvalidDataException($"The author index for comment '{address}' is out of range.");

            var cell = worksheet.Cell(address) ?? throw new InvalidDataException($"Invalid comment address '{address}'.");
            var comment = cell.CreateComment(getShapeId(index));
            comment.Author = authors[(int)author];
            worksheet.Workbook.ShapeIdManager.Add(comment.ShapeId);

            if (textReader.ParseCtRst("text", ns) is { IsSuccess: true } text)
            {
                if (text.Value.TryPickT0(out var plainText, out var richText))
                {
                    if (plainText.Length != 0)
                        comment.AddText(plainText.FixNewLines());
                }
                else
                {
                    foreach (var run in richText.Runs)
                        comment.AddText(richText.GetRunText(run).FixNewLines(), run.Dxf);
                }
            }

            while (tree.TryOpen("extLst", ns))
                tree.Skip("extLst");
            tree.Close("comment", ns);
            applyShape(index++, comment);
        }

        tree.Close("commentList", ns);
        if (tree.TryOpen("extLst", ns))
            tree.Skip("extLst");
        tree.Close("comments", ns);
    }
}
