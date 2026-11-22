using System.Xml;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using static Aid.Xlsx.Model.Workbooks.XlsxWorkbook;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.SpreadsheetML.Strings;

internal class SharedStringTableWriter
{
    internal static void GenerateSharedStringTablePartContent(XlsxWorkbook workbook, Stream partStream,
        SaveContext context)
    {
        // Call all table headers to make sure their names are filled
        workbook.Worksheets.ForEach(w => w.Tables.ForEach(t => _ = ((XlsxTable)t).FieldNames.Count));

        var settings = new XmlWriterSettings
        {
            Encoding = XlsxHelper.NoBomUTF8
        };
        using var xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();

        // Due to streaming and XlsxWorkbook structure, we don't know count before strings are written.
        // Attributes count and uniqueCount are optional thus are omitted.
        xml.WriteStartElement("x", "sst", Main2006SsNs);

        var sst = workbook.SharedStringTable;
        var map = sst.GetConsecutiveMap();
        context.SstMap = map;
        for (var sharedStringId = 0; sharedStringId < map.Count; ++sharedStringId)
        {
            var continuousId = map[sharedStringId];
            if (continuousId < 0)
                continue;

            var richText = sst.GetRichText(sharedStringId);
            if (richText is not null)
            {
                xml.WriteStartElement("si", Main2006SsNs);
                TextSerializer.WriteRichTextElements(xml, richText, context);
                xml.WriteEndElement(); // si
            }
            else
            {
                xml.WriteStartElement("si", Main2006SsNs);
                xml.WriteStartElement("t", Main2006SsNs);
                var sharedString = sst[sharedStringId];
                if (!sharedString.Trim().Equals(sharedString))
                    xml.WritePreserveSpaceAttr();

                xml.WriteString(XmlEncoder.EncodeString(sharedString));
                xml.WriteEndElement(); // t
                xml.WriteEndElement(); // si
            }
        }

        xml.WriteEndElement(); // SharedStringTable
        xml.Close();
    }
}
