using System.Xml;

namespace Aid.Xlsx.Opc.Xml;

internal static class XmlReaderExtension
{
    public static LineInfo GetLineInfo(this XmlReader reader)
    {
        if (reader is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            return new LineInfo(lineInfo.LineNumber, lineInfo.LinePosition);

        return new LineInfo(0, 0);
    }
}
