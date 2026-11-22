using System.Xml;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.PivotTables;
using Aid.Xlsx.Utils;
using static Aid.Xlsx.Internal.IO.OpenXmlConst;

namespace Aid.Xlsx.Internal.IO
{
    internal class PivotCacheRecordsWriter
    {
        internal static void WriteContent(Stream partStream, XlsxPivotCache pivotCache)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = XlsxHelper.NoBomUTF8
            };

            using var xml = XmlWriter.Create(partStream, settings);

            xml.WriteStartDocument();
            xml.WriteStartElement("pivotCacheRecords", Main2006SsNs);
            xml.WriteAttributeString("xmlns", "r", null, RelationshipsNs);
            xml.WriteAttributeString("xmlns", "mc", null, MarkupCompatibilityNs);

            // Mark revision as ignorable extension
            xml.WriteAttributeString("mc", "Ignorable", null, "xr");
            xml.WriteAttributeString("xmlns", "xr", null, RevisionNs);

            var recordCount = pivotCache.RecordCount;
            var fieldCount = pivotCache.FieldCount;
            for (var recordIdx = 0; recordIdx < recordCount; ++recordIdx)
            {
                xml.WriteStartElement("r");
                for (var fieldIdx = 0; fieldIdx < fieldCount; ++fieldIdx)
                {
                    var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                    var value = fieldValues.GetValue(recordIdx);
                    switch (value.Type)
                    {
                        case XlsxPivotCacheValueType.Missing:
                            xml.WriteEmptyElement("m");
                            break;
                        case XlsxPivotCacheValueType.Number:
                            xml.WriteStartElement("n");
                            xml.WriteAttribute("v", value.GetNumber());
                            xml.WriteEndElement();
                            break;
                        case XlsxPivotCacheValueType.Boolean:
                            xml.WriteStartElement("b");
                            xml.WriteAttribute("v", value.GetBoolean());
                            xml.WriteEndElement();
                            break;
                        case XlsxPivotCacheValueType.Error:
                            xml.WriteStartElement("b");
                            xml.WriteAttribute("v", value.GetError().ToDisplayString());
                            xml.WriteEndElement();
                            break;
                        case XlsxPivotCacheValueType.String:
                            xml.WriteStartElement("s");
                            xml.WriteAttribute("v", fieldValues.GetText(value));
                            xml.WriteEndElement();
                            break;
                        case XlsxPivotCacheValueType.DateTime:
                            xml.WriteStartElement("d");
                            xml.WriteAttribute("v", value.GetDateTime());
                            xml.WriteEndElement();
                            break;
                        case XlsxPivotCacheValueType.Index:
                            xml.WriteStartElement("x");
                            xml.WriteAttribute("v", value.GetIndex());
                            xml.WriteEndElement();
                            break;
                        default:
                            throw new NotSupportedException();
                    }
                }
                xml.WriteEndElement(); // r
            }

            xml.WriteEndElement(); // pivotCacheRecords
        }
    }
}
