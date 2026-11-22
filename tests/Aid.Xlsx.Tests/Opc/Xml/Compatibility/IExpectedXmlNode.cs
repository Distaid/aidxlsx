#nullable enable
using Aid.Xlsx.Opc.Xml;

namespace Aid.Xlsx.Tests.Opc.Xml.Compatibility;

internal interface IExpectedXmlNode
{
    void AssertMatches(IXmlReader reader);
}
