// Keep this file CodeMaid organised and cleaned
using System.Xml.Linq;
using System.Xml;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class XDocumentExtensions
{
    public static XDocument? Load(Stream stream)
    {
        using var reader = XmlReader.Create(stream);
        try
        {
            return XDocument.Load(reader);
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
