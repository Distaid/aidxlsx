#nullable enable
using Aid.Xlsx.Opc.Xml;

namespace Aid.Xlsx.Tests.IO;

internal interface IExpectedXmlNode
{
    void AssertMatches(IXmlReader reader);
}
